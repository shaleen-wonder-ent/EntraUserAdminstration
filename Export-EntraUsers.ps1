#Requires -Version 5.1
<#
.SYNOPSIS
    Entra ID Comprehensive User Export Tool for IdP Migration

.DESCRIPTION
    Exports detailed user data from Microsoft Entra ID (formerly Azure AD) for
    migration to another Identity Provider.  Supports multiple export scopes and
    collects every attribute needed for a complete identity migration.

    DATA COLLECTED PER USER
    ------------------------
      * Full profile (50+ attributes incl. on-premises sync info, ImmutableId)
      * Transitive group memberships  (Security / M365 / Distribution, with AD info)
      * Entra ID directory role assignments  (direct + PIM eligible/active)
      * Azure Resource RBAC  (all subscriptions, optional)
      * Enterprise app role assignments + OAuth2 delegated permission grants
      * MFA / authentication methods  (Authenticator, Phone, FIDO2, OATH, etc.)
      * Assigned licenses (by SKU)
      * Registered & joined devices
      * Manager and direct reports
      * Sign-in activity (last interactive + non-interactive timestamps)
      * Extension / custom attributes

    EXPORT SCOPES
    -------------
      Tenant             -> all member users in the directory
      SecurityGroup      -> members of a security group (transitive)
      M365Group          -> members of a Microsoft 365 group
      Department         -> users filtered by Department attribute
      AdministrativeUnit -> members of an Administrative Unit
      SpecificUsers      -> comma-separated UPNs / Object IDs

    OUTPUT FILES
    ------------
      users_full.json             - complete data for all users
      per_user/<upn>.json         - one JSON file per user
      users_profile.csv           - flat profile attributes
      users_groups.csv            - group memberships
      users_directory_roles.csv   - Entra ID directory roles
      users_azure_rbac.csv        - Azure resource RBAC  (if enabled)
      users_app_assignments.csv   - enterprise application role assignments
      users_oauth2_grants.csv     - OAuth2 delegated permission grants
      users_auth_methods.csv      - MFA and authentication methods
      users_licenses.csv          - license assignments
      users_devices.csv           - registered and joined devices
      export_summary.txt          - run statistics and manifest
      export.log                  - detailed run log

.PARAMETER OutputDir
    Path where all output files are written.
    Defaults to .\EntraExport_<yyyyMMdd_HHmmss> beside the script.

.PARAMETER Scope
    Export scope. If omitted an interactive menu is shown.
    Valid: Tenant | SecurityGroup | M365Group | Department | AdministrativeUnit | SpecificUsers

.PARAMETER ScopeValue
    Name or Object ID of the target group / department / AU (not used for Tenant).

.PARAMETER ExportFormat
    Output format: JSON | CSV | Both  (default: Both)

.PARAMETER IncludeAzureRBAC
    Also query Azure Resource Manager for RBAC assignments across all accessible
    subscriptions. Requires Az.Accounts and Az.Resources modules.

.PARAMETER SkipAuthMethods
    Skip MFA / authentication methods (requires UserAuthenticationMethod.Read.All).

.PARAMETER SkipDevices
    Skip device registrations.

.PARAMETER IncludeGuests
    Also include Guest users (default: Members only).

.EXAMPLE
    # Interactive mode (recommended for first run)
    .\Export-EntraUsers.ps1

.EXAMPLE
    # Dump entire tenant, both formats, including Azure RBAC
    .\Export-EntraUsers.ps1 -Scope Tenant -ExportFormat Both -IncludeAzureRBAC

.EXAMPLE
    # Dump one security group
    .\Export-EntraUsers.ps1 -Scope SecurityGroup -ScopeValue "SG-Finance-All"

.EXAMPLE
    # Export two specific users, JSON only
    .\Export-EntraUsers.ps1 -Scope SpecificUsers -ScopeValue "alice@contoso.com,bob@contoso.com" -ExportFormat JSON

.NOTES
    Required Graph API permissions (delegated OR app-only):
        User.Read.All, Group.Read.All, GroupMember.Read.All,
        Directory.Read.All, RoleManagement.Read.All,
        RoleManagement.Read.Directory, AuditLog.Read.All, Application.Read.All
    Optional:
        UserAuthenticationMethod.Read.All  (if SkipAuthMethods is not set)
    Optional for Azure RBAC:
        Reader role on target subscriptions

    Required PowerShell modules (auto-installed if missing):
        Microsoft.Graph.* family, Az.Accounts, Az.Resources (optional)
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]
    $OutputDir,

    [ValidateSet('Tenant','SecurityGroup','M365Group','Department',
                 'AdministrativeUnit','SpecificUsers')]
    [string]
    $Scope,

    [string]
    $ScopeValue,

    # Your Entra ID tenant ID or primary domain, e.g. 'contoso.onmicrosoft.com' or
    # the GUID from https://entra.microsoft.com -> Overview. Supplying this forces
    # sign-in against your work/school tenant and avoids the MSA consumer error.
    [string]
    $TenantId,

    [ValidateSet('JSON','CSV','Both')]
    [string]
    $ExportFormat = 'Both',

    [switch]
    $IncludeAzureRBAC,

    [switch]
    $SkipAuthMethods,

    [switch]
    $SkipDevices,

    [switch]
    $IncludeGuests
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# SCRIPT-LEVEL STATE
# ---------------------------------------------------------------------------

$Script:Version   = '2.0.0'
$Script:StartTime = Get-Date

$Script:TenantId          = ''
$Script:TenantName        = ''
$Script:ScopeType         = ''
$Script:ScopeValue        = ''
$Script:ExportFormat      = $ExportFormat
$Script:IncludeAzureRBAC  = $IncludeAzureRBAC.IsPresent
$Script:SkipAuthMethods   = $SkipAuthMethods.IsPresent
$Script:SkipDevices       = $SkipDevices.IsPresent
$Script:IncludeGuests     = $IncludeGuests.IsPresent
$Script:LogFile           = ''
$Script:OutputDir         = if ($OutputDir) {
    $OutputDir
} else {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
    Join-Path $scriptDir "EntraExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
}

$Script:Stats = @{
    TotalUsers     = 0
    UsersProcessed = 0
    Errors         = 0
    SkippedUsers   = 0
}

# Base Graph scopes; auth-methods scope added conditionally after options menu
$Script:GraphScopes = @(
    'User.Read.All'
    'Group.Read.All'
    'GroupMember.Read.All'
    'Directory.Read.All'
    'RoleManagement.Read.All'
    'RoleManagement.Read.Directory'
    'AuditLog.Read.All'
    'Application.Read.All'
)

$Script:RequiredGraphModules = @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Identity.DirectoryManagement'
    'Microsoft.Graph.Identity.Governance'
    'Microsoft.Graph.Applications'
    'Microsoft.Graph.Identity.SignIns'
)

$Script:RequiredAzModules = @(
    'Az.Accounts'
    'Az.Resources'
)

# Tenant ID collected interactively (when -TenantId param is not provided)
$Script:InputTenantId = ''

# All user properties requested in a single Graph call
$Script:UserSelectProps = (
    'id,userPrincipalName,displayName,givenName,surname,mailNickname,' +
    'mail,proxyAddresses,otherMails,mobilePhone,businessPhones,' +
    'jobTitle,department,companyName,employeeId,employeeType,' +
    'accountEnabled,userType,createdDateTime,lastPasswordChangeDateTime,' +
    'passwordPolicies,' +
    'usageLocation,preferredLanguage,preferredDataLocation,' +
    'city,state,country,streetAddress,postalCode,officeLocation,' +
    'onPremisesSyncEnabled,onPremisesImmutableId,onPremisesDomainName,' +
    'onPremisesSamAccountName,onPremisesUserPrincipalName,' +
    'onPremisesLastSyncDateTime,onPremisesDistinguishedName,' +
    'onPremisesProvisioningErrors,' +
    'externalUserState,externalUserStateChangeDateTime,' +
    'signInSessionsValidFromDateTime,refreshTokensValidFromDateTime,' +
    'assignedLicenses,assignedPlans,' +
    'aboutMe,mySite,interests,pastProjects,responsibilities,skills,' +
    'showInAddressList,signInActivity'
)

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)][string] $Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')]
        [string] $Level = 'INFO',
        [switch] $NoConsole
    )
    $ts      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$ts][$Level] $Message"

    if (-not $NoConsole) {
        $color = switch ($Level) {
            'ERROR'   { 'Red'      }
            'WARN'    { 'Yellow'   }
            'SUCCESS' { 'Green'    }
            'DEBUG'   { 'DarkGray' }
            default   { 'White'    }
        }
        Write-Host $logLine -ForegroundColor $color
    }
    if ($Script:LogFile -and (Test-Path (Split-Path $Script:LogFile -Parent))) {
        try {
            $fs = [System.IO.File]::Open(
                $Script:LogFile,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $sw = [System.IO.StreamWriter]::new($fs, [System.Text.Encoding]::UTF8)
            $sw.WriteLine($logLine)
            $sw.Dispose()
            $fs.Dispose()
        } catch { <# silently skip log write if still locked #> }
    }
}

function Write-Banner {
    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host "  Entra ID Comprehensive User Export Tool  v$($Script:Version)" -ForegroundColor Cyan
    Write-Host "  IdP Migration Export  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor DarkCyan
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host ''
}

# ---------------------------------------------------------------------------
# MODULE MANAGEMENT
# ---------------------------------------------------------------------------

function Install-RequiredModules {
    param(
        [Parameter(Mandatory)][string[]] $Modules,
        [string] $Label = 'modules'
    )
    Write-Log "Checking $Label ..."
    foreach ($mod in $Modules) {
        if (-not (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue)) {
            Write-Log "  Installing $mod from PSGallery ..." 'WARN'
            try {
                Install-Module -Name $mod -Scope CurrentUser -Force `
                               -AllowClobber -Repository PSGallery -ErrorAction Stop
                Write-Log "  $mod installed." 'SUCCESS'
            } catch {
                Write-Log "  Failed to install ${mod}: $($_.Exception.Message)" 'ERROR'
                throw
            }
        }
        # Az.Resources 7.x/8.x on Windows PowerShell 5.1 emits TypeInitializationException
        # warnings from its AutoRest sub-modules during import; these are non-fatal and the
        # module loads correctly. Redirect all streams to suppress the noise.
        if ($mod -like 'Az.*') {
            $null = Import-Module $mod -DisableNameChecking -Force -ErrorAction Stop *>&1 |
                    Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] -or
                                   $_.FullyQualifiedErrorId -notmatch 'RegisterAzModule' }
        } else {
            Import-Module $mod -DisableNameChecking -Force -ErrorAction Stop
        }
        Write-Log "  $mod loaded." 'DEBUG'
    }
}

# ---------------------------------------------------------------------------
# CONNECTION
# ---------------------------------------------------------------------------

# The MSA consumer tenant ID — accounts that land here cannot use the
# organisation / directory APIs that this script requires.
$Script:MsaConsumerTenantId = '9188040d-6c67-4c5b-b112-36a304b66dad'

function Connect-ToEntra {
    Write-Log 'Connecting to Microsoft Graph ...'
    try {
        $connectParams = @{
            Scopes    = $Script:GraphScopes
            NoWelcome = $true
            ErrorAction = 'Stop'
        }
        # Resolve tenant ID: CLI param takes precedence, then interactively collected value.
        $resolvedTenantId = if ($TenantId) { $TenantId } else { $Script:InputTenantId }
        if ($resolvedTenantId) {
            $connectParams['TenantId'] = $resolvedTenantId
            Write-Log "  Scoping sign-in to tenant: $resolvedTenantId"
        } else {
            Write-Log '  No TenantId supplied - the sign-in dialog will show all accounts.' 'WARN'
            Write-Log '  If you accidentally pick a personal (@outlook.com/@hotmail.com) account the script will stop with a clear error.' 'WARN'
        }

        Connect-MgGraph @connectParams
        $ctx = Get-MgContext

        # Detect MSA consumer sign-in before calling any tenant-only API.
        # The consumer tenant GUID is a well-known constant for personal accounts.
        if ($ctx.TenantId -eq $Script:MsaConsumerTenantId) {
            Write-Host ''
            Write-Host '  ERROR: You signed in with a personal Microsoft account (MSA).' -ForegroundColor Red
            Write-Host '  This script requires a WORK or SCHOOL account that belongs to an Entra ID tenant.' -ForegroundColor Red
            Write-Host ''
            Write-Host '  How to fix:' -ForegroundColor Yellow
            Write-Host '  1. Re-run the script and sign in with your organisation account (e.g. you@contoso.com).' -ForegroundColor Yellow
            Write-Host '  2. Or supply your tenant ID explicitly to bypass the account picker:' -ForegroundColor Yellow
            Write-Host '       .\Export-EntraUsers.ps1 -TenantId "contoso.onmicrosoft.com"' -ForegroundColor Cyan
            Write-Host '       .\Export-EntraUsers.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"' -ForegroundColor Cyan
            Write-Host '  (Your tenant ID is visible in Entra admin center -> Overview -> Tenant ID)' -ForegroundColor Yellow
            Write-Host ''
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            throw 'MSA consumer account detected. A work/school account is required.'
        }

        $org = (Get-MgOrganization -ErrorAction Stop) | Select-Object -First 1

        $Script:TenantId   = $ctx.TenantId
        $Script:TenantName = $org.DisplayName

        Write-Log "Connected -> Tenant: $($Script:TenantName) [$($Script:TenantId)]" 'SUCCESS'
    } catch {
        $msg = $_.Exception.Message
        # Surface a friendlier message for the known MSA BadRequest error in case
        # the tenant-ID check above is bypassed somehow.
        if ($msg -match 'MSA accounts' -or $msg -match 'DirectoryServices') {
            Write-Host ''
            Write-Host '  ERROR: The API returned "not supported for MSA accounts".' -ForegroundColor Red
            Write-Host '  You must sign in with a work/school Entra ID account, not a personal @outlook.com / @hotmail.com account.' -ForegroundColor Red
            Write-Host '  Retry with: .\Export-EntraUsers.ps1 -TenantId "your-tenant.onmicrosoft.com"' -ForegroundColor Cyan
            Write-Host ''
        }
        Write-Log "Graph connection failed: $msg" 'ERROR'
        throw
    }
}

function Connect-ToAzure {
    Write-Log 'Connecting to Azure Resource Manager (for RBAC) ...'
    try {
        # $Script:TenantId is always set by Connect-ToEntra before this is called
        $null = Connect-AzAccount -Tenant $Script:TenantId -ErrorAction Stop
        Write-Log 'Azure RM connected.' 'SUCCESS'
    } catch {
        # InteractiveBrowserCredential can fail due to Azure.Identity assembly version
        # conflicts when Microsoft.Graph is loaded in the same session. Fall back to
        # device code flow which avoids that code path entirely.
        try {
            Write-Log 'Interactive browser auth failed, retrying with device code ...' 'WARN'
            $null = Connect-AzAccount -Tenant $Script:TenantId -UseDeviceAuthentication -ErrorAction Stop
            Write-Log 'Azure RM connected (device code).' 'SUCCESS'
        } catch {
            Write-Log "Azure RM connection failed - Azure RBAC will be skipped. $($_.Exception.Message)" 'WARN'
            $Script:IncludeAzureRBAC = $false
        }
    }
}

# ---------------------------------------------------------------------------
# INTERACTIVE MENUS
# ---------------------------------------------------------------------------

function Show-ScopeMenu {
    Write-Host ''
    Write-Host ('  ' + ('=' * 56)) -ForegroundColor Cyan
    Write-Host '  SELECT EXPORT SCOPE' -ForegroundColor Cyan
    Write-Host ('  ' + ('=' * 56)) -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  [1]  Entire Tenant           - all member users' -ForegroundColor White
    Write-Host '  [2]  Security Group          - by name or Object ID' -ForegroundColor White
    Write-Host '  [3]  Microsoft 365 Group     - by name or Object ID' -ForegroundColor White
    Write-Host '  [4]  Department              - exact department name' -ForegroundColor White
    Write-Host '  [5]  Administrative Unit     - by name or Object ID' -ForegroundColor White
    Write-Host '  [6]  Specific User(s)        - comma-separated UPNs or Object IDs' -ForegroundColor White
    Write-Host ''

    $choice = Read-Host '  Choice [1-6]'
    switch ($choice.Trim()) {
        '1' { return @{ Scope = 'Tenant';             Value = $null } }
        '2' {
            $v = Read-Host '  Security Group display name or Object ID'
            return @{ Scope = 'SecurityGroup';      Value = $v.Trim() }
        }
        '3' {
            $v = Read-Host '  M365 Group display name or Object ID'
            return @{ Scope = 'M365Group';           Value = $v.Trim() }
        }
        '4' {
            $v = Read-Host '  Department name (exact, case-sensitive)'
            return @{ Scope = 'Department';          Value = $v.Trim() }
        }
        '5' {
            $v = Read-Host '  Administrative Unit name or Object ID'
            return @{ Scope = 'AdministrativeUnit';  Value = $v.Trim() }
        }
        '6' {
            $v = Read-Host '  UPN(s) or Object ID(s), comma-separated'
            return @{ Scope = 'SpecificUsers';       Value = $v.Trim() }
        }
        default {
            Write-Log 'Invalid choice - defaulting to Tenant scope.' 'WARN'
            return @{ Scope = 'Tenant'; Value = $null }
        }
    }
}

function Show-OptionsMenu {
    Write-Host ''
    Write-Host ('  ' + ('=' * 56)) -ForegroundColor Cyan
    Write-Host '  EXPORT OPTIONS' -ForegroundColor Cyan
    Write-Host ('  ' + ('=' * 56)) -ForegroundColor Cyan
    Write-Host ''

    # --- Tenant ID -------------------------------------------------------
    Write-Host '  Tenant ID (recommended if you have personal MS accounts in your browser)' -ForegroundColor White
    Write-Host '  Leave blank to use the default account picker.' -ForegroundColor DarkGray
    Write-Host '  Find your Tenant ID at: https://entra.microsoft.com -> Overview -> Tenant ID' -ForegroundColor DarkGray
    $tenantAns = Read-Host '  Tenant ID or domain [e.g. contoso.onmicrosoft.com, press Enter to skip]'
    $Script:InputTenantId = $tenantAns.Trim()
    Write-Host ''

    $guestAns = Read-Host '  Include Guest users? [Y/n, default Y]'
    $Script:IncludeGuests = (-not ($guestAns -match '^[Nn]'))

    $rbacAns = Read-Host '  Include Azure Resource RBAC? (requires Az modules + separate sign-in) [Y/n, default Y]'
    $Script:IncludeAzureRBAC = (-not ($rbacAns -match '^[Nn]'))

    $mfaAns = Read-Host '  Include MFA / auth methods? (UserAuthenticationMethod.Read.All) [Y/n, default Y]'
    $Script:SkipAuthMethods = ($mfaAns -match '^[Nn]')

    $devAns = Read-Host '  Include registered devices? [Y/n, default Y]'
    $Script:SkipDevices = ($devAns -match '^[Nn]')

    Write-Host ''
    Write-Host '  Output format:' -ForegroundColor White
    Write-Host '  [1] Both JSON and CSV (default)' -ForegroundColor White
    Write-Host '  [2] JSON only' -ForegroundColor White
    Write-Host '  [3] CSV only' -ForegroundColor White
    $fmtAns = Read-Host '  Choice [1-3]'
    $Script:ExportFormat = switch ($fmtAns.Trim()) {
        '2' { 'JSON' }
        '3' { 'CSV'  }
        default { 'Both' }
    }
    Write-Host ''
}

# ---------------------------------------------------------------------------
# SCOPE -> USER IDs
# ---------------------------------------------------------------------------

$Script:GuidRegex = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

function Get-UserIdsByScope {
    param(
        [Parameter(Mandatory)][string] $ScopeType,
        [string] $ScopeVal
    )
    Write-Log "Resolving users for scope '$ScopeType'$(if ($ScopeVal) { " -> '$ScopeVal'" }) ..."

    $typeFilter = if ($Script:IncludeGuests) { '' } else { " and userType eq 'Member'" }

    switch ($ScopeType) {

        'Tenant' {
            if ($typeFilter) {
                $filter = $typeFilter -replace '^\s*and\s+', ''
                $users = Get-MgUser -All -Filter $filter -Property 'id' -ErrorAction Stop
            } else {
                $users = Get-MgUser -All -Property 'id' -ErrorAction Stop
            }
            return @($users | Select-Object -ExpandProperty Id)
        }

        { $_ -in 'SecurityGroup','M365Group' } {
            $group = $null
            if ($ScopeVal -match $Script:GuidRegex) {
                $group = Get-MgGroup -GroupId $ScopeVal -Property 'id,displayName' -ErrorAction Stop
            } else {
                $esc   = $ScopeVal -replace "'","''"
                $group = Get-MgGroup -Filter "displayName eq '$esc'" -Property 'id,displayName' -ErrorAction Stop |
                         Select-Object -First 1
            }
            if (-not $group) { throw "Group not found: $ScopeVal" }
            Write-Log "Found group: $($group.DisplayName) [$($group.Id)]" 'SUCCESS'

            $members = Get-MgGroupTransitiveMember -GroupId $group.Id -All -ErrorAction Stop
            return @($members |
                     Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } |
                     Select-Object -ExpandProperty Id)
        }

        'Department' {
            $esc    = $ScopeVal -replace "'","''"
            $filter = "department eq '$esc'$typeFilter"
            $filter = $filter -replace '^\s*and\s+', ''
            $users  = Get-MgUser -All -Filter $filter -Property 'id' -ErrorAction Stop
            if (-not $users) { Write-Log "No users found in department: $ScopeVal" 'WARN' }
            return @($users | Select-Object -ExpandProperty Id)
        }

        'AdministrativeUnit' {
            $au = $null
            if ($ScopeVal -match $Script:GuidRegex) {
                $au = Get-MgDirectoryAdministrativeUnit -AdministrativeUnitId $ScopeVal -ErrorAction Stop
            } else {
                $esc = $ScopeVal -replace "'","''"
                $au  = Get-MgDirectoryAdministrativeUnit -Filter "displayName eq '$esc'" -ErrorAction Stop |
                       Select-Object -First 1
            }
            if (-not $au) { throw "Administrative Unit not found: $ScopeVal" }
            Write-Log "Found AU: $($au.DisplayName) [$($au.Id)]" 'SUCCESS'

            $members = Get-MgDirectoryAdministrativeUnitMember -AdministrativeUnitId $au.Id -All -ErrorAction Stop
            return @($members |
                     Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } |
                     Select-Object -ExpandProperty Id)
        }

        'SpecificUsers' {
            $ids = [System.Collections.Generic.List[string]]::new()
            foreach ($entry in ($ScopeVal -split ',')) {
                $entry = $entry.Trim()
                if (-not $entry) { continue }
                if ($entry -match $Script:GuidRegex) {
                    $ids.Add($entry)
                } else {
                    try {
                        $u = Get-MgUser -UserId $entry -Property 'id' -ErrorAction Stop
                        $ids.Add($u.Id)
                    } catch {
                        Write-Log "User not found (skipping): $entry" 'WARN'
                        $Script:Stats.SkippedUsers++
                    }
                }
            }
            return $ids.ToArray()
        }
    }
}

# ---------------------------------------------------------------------------
# PER-USER DATA COLLECTORS
# ---------------------------------------------------------------------------

function Get-FullUserProfile {
    param([Parameter(Mandatory)][string] $UserId)
    try {
        $user = Get-MgUser -UserId $UserId -Property $Script:UserSelectProps -ErrorAction Stop

        # Extension attributes (directory schema extensions + open extensions)
        $extAttrs = [ordered]@{}
        try {
            $extUser = Get-MgUser -UserId $UserId -ExpandProperty 'extensions' -ErrorAction SilentlyContinue
            foreach ($e in $extUser.Extensions) {
                $extAttrs[$e.Id] = $e.AdditionalProperties
            }
        } catch { }

        return [PSCustomObject]@{ User = $user; ExtensionAttrs = $extAttrs }
    } catch {
        Write-Log "  [Profile] Error for ${UserId}: $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Get-UserGroupMemberships {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    try {
        $memberships = Get-MgUserTransitiveMemberOf -UserId $UserId -All -ErrorAction Stop
        $result = [System.Collections.Generic.List[hashtable]]::new()

        foreach ($m in $memberships) {
            $entry = $null
            switch ($m.AdditionalProperties['@odata.type']) {

                '#microsoft.graph.group' {
                    try {
                        $g = Get-MgGroup -GroupId $m.Id `
                             -Property 'id,displayName,description,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,onPremisesNetBiosName,onPremisesDomainName,onPremisesSamAccountName,onPremisesDistinguishedName,membershipRule,membershipRuleProcessingState' `
                             -ErrorAction Stop
                        $cat = if ($g.GroupTypes -contains 'Unified') { 'M365Group' }
                               elseif ($g.SecurityEnabled -and $g.GroupTypes.Count -eq 0) { 'SecurityGroup' }
                               elseif ((-not $g.SecurityEnabled) -and $g.MailEnabled) { 'DistributionList' }
                               else { 'Other' }
                        $entry = @{
                            ObjectType               = 'Group'
                            Id                       = $g.Id
                            DisplayName              = $g.DisplayName
                            Description              = $g.Description
                            Category                 = $cat
                            SecurityEnabled          = $g.SecurityEnabled
                            MailEnabled              = $g.MailEnabled
                            IsDynamic                = ($g.GroupTypes -contains 'DynamicMembership')
                            MembershipRule           = $g.MembershipRule
                            OnPremisesSynced         = $g.OnPremisesSyncEnabled
                            OnPremisesNetBiosName    = $g.OnPremisesNetBiosName
                            OnPremisesDomainName     = $g.OnPremisesDomainName
                            OnPremisesSamAccountName = $g.OnPremisesSamAccountName
                            OnPremisesDN             = $g.OnPremisesDistinguishedName
                        }
                    } catch {
                        $entry = @{
                            ObjectType  = 'Group'
                            Id          = $m.Id
                            DisplayName = '(access error)'
                            Category    = 'Unknown'
                        }
                    }
                }

                '#microsoft.graph.directoryRole' {
                    $entry = @{
                        ObjectType  = 'DirectoryRole'
                        Id          = $m.Id
                        DisplayName = $m.AdditionalProperties['displayName']
                        Category    = 'DirectoryRole'
                    }
                }

                '#microsoft.graph.administrativeUnit' {
                    $entry = @{
                        ObjectType  = 'AdministrativeUnit'
                        Id          = $m.Id
                        DisplayName = $m.AdditionalProperties['displayName']
                        Category    = 'AdministrativeUnit'
                    }
                }

                default {
                    $entry = @{
                        ObjectType = 'Other'
                        Id         = $m.Id
                        OdataType  = $m.AdditionalProperties['@odata.type']
                    }
                }
            }
            if ($entry) { $result.Add($entry) }
        }
        return $result.ToArray()
    } catch {
        Write-Log "  [Groups] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Get-UserDirectoryRoles {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    $result = [System.Collections.Generic.List[hashtable]]::new()

    # PIM and directly-assigned roles via unified RBAC endpoint
    try {
        $assignments = Get-MgRoleManagementDirectoryRoleAssignment `
                           -Filter "principalId eq '$UserId'" -All `
                           -ExpandProperty 'roleDefinition' -ErrorAction Stop
        foreach ($a in $assignments) {
            $result.Add(@{
                AssignmentId     = $a.Id
                RoleDefinitionId = $a.RoleDefinitionId
                RoleDisplayName  = $a.RoleDefinition.DisplayName
                RoleDescription  = $a.RoleDefinition.Description
                DirectoryScopeId = $a.DirectoryScopeId
                AssignmentType   = 'Assigned'
                Source           = 'UnifiedRBAC'
            })
        }
    } catch {
        Write-Log "  [DirRoles-RBAC] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
    }

    # Supplement via transitive membership (catches legacy role objects)
    try {
        $roleMembers = Get-MgUserTransitiveMemberOf -UserId $UserId -All -ErrorAction SilentlyContinue
        foreach ($r in ($roleMembers | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' })) {
            $roleDef = Get-MgDirectoryRole -DirectoryRoleId $r.Id -ErrorAction SilentlyContinue
            if ($roleDef) {
                $alreadyHave = $result | Where-Object { $_.RoleDefinitionId -eq $roleDef.RoleTemplateId }
                if (-not $alreadyHave) {
                    $result.Add(@{
                        RoleDefinitionId = $roleDef.RoleTemplateId
                        RoleDisplayName  = $roleDef.DisplayName
                        RoleDescription  = $roleDef.Description
                        DirectoryScopeId = '/'
                        AssignmentType   = 'Assigned'
                        Source           = 'DirectoryRole'
                    })
                }
            }
        }
    } catch { }

    return $result.ToArray()
}

function Get-UserAzureRBAC {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    if (-not $Script:IncludeAzureRBAC) { return @() }
    try {
        $assignments = Get-AzRoleAssignment -ObjectId $UserId -ErrorAction Stop
        return @($assignments | ForEach-Object {
            @{
                RoleDefinitionName = $_.RoleDefinitionName
                RoleDefinitionId   = $_.RoleDefinitionId
                Scope              = $_.Scope
                ResourceGroupName  = $_.ResourceGroupName
                ResourceName       = $_.ResourceName
                ResourceType       = $_.ResourceType
                AssignmentId       = $_.RoleAssignmentId
                CanDelegate        = $_.CanDelegate
                Description        = $_.Description
            }
        })
    } catch {
        Write-Log "  [AzureRBAC] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Get-UserAppAssignments {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    $appRoles = [System.Collections.Generic.List[hashtable]]::new()
    $oauth2   = [System.Collections.Generic.List[hashtable]]::new()

    # Enterprise app role assignments
    try {
        $assignments = Get-MgUserAppRoleAssignment -UserId $UserId -All -ErrorAction Stop
        foreach ($a in $assignments) {
            $roleName = $a.AppRoleId
            if ($a.AppRoleId -eq '00000000-0000-0000-0000-000000000000') {
                $roleName = 'Default Access'
            } else {
                try {
                    $sp = Get-MgServicePrincipal -ServicePrincipalId $a.ResourceId `
                              -Property 'appRoles' -ErrorAction SilentlyContinue
                    if ($sp) {
                        $match = $sp.AppRoles | Where-Object { $_.Id -eq $a.AppRoleId } |
                                 Select-Object -First 1
                        if ($match) { $roleName = $match.DisplayName }
                    }
                } catch { }
            }
            $appRoles.Add(@{
                ApplicationName    = $a.ResourceDisplayName
                ServicePrincipalId = $a.ResourceId
                AppRoleId          = $a.AppRoleId
                AppRoleName        = $roleName
                CreatedDateTime    = $a.CreatedDateTime
                AssignmentId       = $a.Id
            })
        }
    } catch {
        Write-Log "  [AppRoles] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
    }

    # OAuth2 delegated permission grants (user consents)
    try {
        $grants = Get-MgUserOauth2PermissionGrant -UserId $UserId -All -ErrorAction Stop
        foreach ($g in $grants) {
            $clientName = $g.ClientId
            try {
                $sp = Get-MgServicePrincipal -ServicePrincipalId $g.ClientId `
                          -Property 'displayName,appId' -ErrorAction SilentlyContinue
                if ($sp) { $clientName = $sp.DisplayName }
            } catch { }
            $oauth2.Add(@{
                ClientAppName = $clientName
                ClientId      = $g.ClientId
                ResourceId    = $g.ResourceId
                Scope         = $g.Scope
                ConsentType   = $g.ConsentType
            })
        }
    } catch {
        Write-Log "  [OAuth2] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
    }

    return @{
        AppRoleAssignments = $appRoles.ToArray()
        OAuth2Grants       = $oauth2.ToArray()
    }
}

function Get-UserAuthMethods {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    if ($Script:SkipAuthMethods) { return @() }
    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $UserId -All -ErrorAction Stop
        return @($methods | ForEach-Object {
            $m = $_
            switch ($m.AdditionalProperties['@odata.type']) {
                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                    @{
                        MethodType      = 'MicrosoftAuthenticator'
                        DisplayName     = $m.AdditionalProperties['displayName']
                        DeviceTag       = $m.AdditionalProperties['deviceTag']
                        PhoneAppVersion = $m.AdditionalProperties['phoneAppVersion']
                        CreatedDateTime = $m.AdditionalProperties['createdDateTime']
                    }
                }
                '#microsoft.graph.phoneAuthenticationMethod' {
                    @{
                        MethodType  = 'Phone'
                        PhoneNumber = $m.AdditionalProperties['phoneNumber']
                        PhoneType   = $m.AdditionalProperties['phoneType']
                    }
                }
                '#microsoft.graph.emailAuthenticationMethod' {
                    @{
                        MethodType   = 'Email'
                        EmailAddress = $m.AdditionalProperties['emailAddress']
                    }
                }
                '#microsoft.graph.fido2AuthenticationMethod' {
                    @{
                        MethodType      = 'FIDO2'
                        DisplayName     = $m.AdditionalProperties['displayName']
                        AaGuid          = $m.AdditionalProperties['aaGuid']
                        Model           = $m.AdditionalProperties['model']
                        CreatedDateTime = $m.AdditionalProperties['createdDateTime']
                    }
                }
                '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' {
                    @{
                        MethodType      = 'WindowsHelloForBusiness'
                        DisplayName     = $m.AdditionalProperties['displayName']
                        CreatedDateTime = $m.AdditionalProperties['createdDateTime']
                    }
                }
                '#microsoft.graph.softwareOathAuthenticationMethod' {
                    @{ MethodType = 'SoftwareOATH' }
                }
                '#microsoft.graph.temporaryAccessPassAuthenticationMethod' {
                    @{
                        MethodType        = 'TemporaryAccessPass'
                        IsUsable          = $m.AdditionalProperties['isUsable']
                        LifetimeInMinutes = $m.AdditionalProperties['lifetimeInMinutes']
                        CreatedDateTime   = $m.AdditionalProperties['createdDateTime']
                    }
                }
                '#microsoft.graph.passwordAuthenticationMethod' {
                    @{
                        MethodType      = 'Password'
                        CreatedDateTime = $m.AdditionalProperties['createdDateTime']
                    }
                }
                default {
                    @{
                        MethodType = 'Unknown'
                        OdataType  = $m.AdditionalProperties['@odata.type']
                    }
                }
            }
        })
    } catch {
        Write-Log "  [AuthMethods] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Get-UserLicenseDetails {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    try {
        $licDetails = Get-MgUserLicenseDetail -UserId $UserId -All -ErrorAction Stop
        return @($licDetails | ForEach-Object {
            @{
                SkuId         = $_.SkuId
                SkuPartNumber = $_.SkuPartNumber
                ServicePlans  = @($_.ServicePlans | ForEach-Object {
                    @{
                        ServicePlanName    = $_.ServicePlanName
                        AppliesTo          = $_.AppliesTo
                        ProvisioningStatus = $_.ProvisioningStatus
                    }
                })
            }
        })
    } catch {
        Write-Log "  [Licenses] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Get-UserDevices {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    if ($Script:SkipDevices) { return @() }
    $result = [System.Collections.Generic.List[hashtable]]::new()

    # Registered devices (BYOD)
    try {
        $registered = Get-MgUserRegisteredDevice -UserId $UserId -All -ErrorAction Stop
        foreach ($item in $registered) {
            try {
                $d = Get-MgDevice -DeviceId $item.Id `
                         -Property 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,isManaged,profileType,enrollmentType,registeredDateTime,approximateLastSignInDateTime,managementType,mdmAppId,onPremisesSyncEnabled,onPremisesLastSyncDateTime' `
                         -ErrorAction Stop
                $result.Add(@{
                    Relationship               = 'Registered'
                    ObjectId                   = $d.Id
                    DeviceId                   = $d.DeviceId
                    DisplayName                = $d.DisplayName
                    OperatingSystem            = $d.OperatingSystem
                    OsVersion                  = $d.OperatingSystemVersion
                    TrustType                  = $d.TrustType
                    IsCompliant                = $d.IsCompliant
                    IsManaged                  = $d.IsManaged
                    ProfileType                = $d.ProfileType
                    EnrollmentType             = $d.EnrollmentType
                    ManagementType             = $d.ManagementType
                    RegisteredDateTime         = $d.RegisteredDateTime
                    ApproximateLastSignIn      = $d.ApproximateLastSignInDateTime
                    OnPremisesSynced           = $d.OnPremisesSyncEnabled
                    OnPremisesLastSyncDateTime = $d.OnPremisesLastSyncDateTime
                })
            } catch {
                $result.Add(@{ Relationship = 'Registered'; ObjectId = $item.Id; Error = 'detail unavailable' })
            }
        }
    } catch {
        Write-Log "  [Devices-registered] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
    }

    # Owned (Entra-joined) devices - add only if not already captured
    try {
        $owned = Get-MgUserOwnedDevice -UserId $UserId -All -ErrorAction Stop
        foreach ($item in $owned) {
            $alreadyCaptured = $result | Where-Object { $_.ObjectId -eq $item.Id }
            if (-not $alreadyCaptured) {
                $result.Add(@{
                    Relationship = 'Owned'
                    ObjectId     = $item.Id
                    DisplayName  = $item.AdditionalProperties['displayName']
                    OdataType    = $item.AdditionalProperties['@odata.type']
                })
            }
        }
    } catch {
        Write-Log "  [Devices-owned] Error for ${UPN}: $($_.Exception.Message)" 'WARN'
    }

    return $result.ToArray()
}

function Get-UserManagerAndReports {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [string] $UPN = $UserId
    )
    $out = @{ Manager = $null; DirectReports = @() }
    try {
        $mgr = Get-MgUserManager -UserId $UserId -ErrorAction SilentlyContinue
        if ($mgr) {
            $out.Manager = @{
                Id                = $mgr.Id
                UserPrincipalName = $mgr.AdditionalProperties['userPrincipalName']
                DisplayName       = $mgr.AdditionalProperties['displayName']
                Mail              = $mgr.AdditionalProperties['mail']
            }
        }
    } catch { }
    try {
        $reports = Get-MgUserDirectReport -UserId $UserId -All -ErrorAction SilentlyContinue
        $out.DirectReports = @($reports | ForEach-Object {
            @{
                Id                = $_.Id
                UserPrincipalName = $_.AdditionalProperties['userPrincipalName']
                DisplayName       = $_.AdditionalProperties['displayName']
            }
        })
    } catch { }
    return $out
}

# ---------------------------------------------------------------------------
# PER-USER ORCHESTRATION
# ---------------------------------------------------------------------------

function Export-SingleUser {
    param(
        [Parameter(Mandatory)][string] $UserId,
        [int] $Index,
        [int] $Total
    )
    $pct = [int](($Index / [Math]::Max($Total, 1)) * 100)
    Write-Progress -Activity 'Exporting Entra Users' `
                   -Status   "[$Index / $Total] Collecting ..." `
                   -PercentComplete $pct

    try {
        # 1 - Profile
        $profileData = Get-FullUserProfile -UserId $UserId
        if ($null -eq $profileData) { $Script:Stats.Errors++; return $null }
        $u   = $profileData.User
        $upn = $u.UserPrincipalName
        Write-Log "  [$Index/$Total] $upn" 'DEBUG'

        # 2 - Groups (transitive)
        $groups      = @(Get-UserGroupMemberships    -UserId $UserId -UPN $upn)

        # 3 - Entra directory roles
        $dirRoles    = @(Get-UserDirectoryRoles      -UserId $UserId -UPN $upn)

        # 4 - Azure Resource RBAC
        $azureRbac   = @(Get-UserAzureRBAC           -UserId $UserId -UPN $upn)

        # 5 - App role assignments + OAuth2 grants
        $apps        = Get-UserAppAssignments      -UserId $UserId -UPN $upn

        # 6 - MFA / auth methods
        $authMethods = @(Get-UserAuthMethods         -UserId $UserId -UPN $upn)

        # 7 - Licenses
        $licenses    = @(Get-UserLicenseDetails      -UserId $UserId -UPN $upn)

        # 8 - Devices
        $devices     = @(Get-UserDevices             -UserId $UserId -UPN $upn)

        # 9 - Manager + direct reports
        $orgInfo     = Get-UserManagerAndReports   -UserId $UserId -UPN $upn

        # 10 - Sign-in activity (comes back inside the user profile object)
        $signIn      = $u.SignInActivity

        $managerUPN  = if ($orgInfo.Manager) { $orgInfo.Manager.UserPrincipalName } else { '' }
        $managerName = if ($orgInfo.Manager) { $orgInfo.Manager.DisplayName }       else { '' }

        $lastSignIn         = if ($signIn) { $signIn.LastSignInDateTime }                  else { $null }
        $lastNonInteractive = if ($signIn) { $signIn.LastNonInteractiveSignInDateTime }    else { $null }

        $Script:Stats.UsersProcessed++

        return [PSCustomObject][ordered]@{
            # --- Identity ---
            ObjectId                        = $u.Id
            UserPrincipalName               = $upn
            DisplayName                     = $u.DisplayName
            GivenName                       = $u.GivenName
            Surname                         = $u.Surname
            MailNickname                    = $u.MailNickname
            Mail                            = $u.Mail
            ProxyAddresses                  = ($u.ProxyAddresses -join ';')
            OtherMails                      = ($u.OtherMails -join ';')
            # --- Contact ---
            MobilePhone                     = $u.MobilePhone
            BusinessPhones                  = ($u.BusinessPhones -join ';')
            # --- Job & Org ---
            JobTitle                        = $u.JobTitle
            Department                      = $u.Department
            CompanyName                     = $u.CompanyName
            EmployeeId                      = $u.EmployeeId
            EmployeeType                    = $u.EmployeeType
            OfficeLocation                  = $u.OfficeLocation
            ManagerUPN                      = $managerUPN
            ManagerDisplayName              = $managerName
            DirectReportCount               = $orgInfo.DirectReports.Count
            # --- Address ---
            StreetAddress                   = $u.StreetAddress
            City                            = $u.City
            State                           = $u.State
            Country                         = $u.Country
            PostalCode                      = $u.PostalCode
            # --- Account state ---
            AccountEnabled                  = $u.AccountEnabled
            UserType                        = $u.UserType
            ExternalUserState               = $u.ExternalUserState
            CreatedDateTime                 = $u.CreatedDateTime
            LastPasswordChangeDateTime      = $u.LastPasswordChangeDateTime
            PasswordPolicies                = $u.PasswordPolicies
            SignInSessionsValidFromDateTime = $u.SignInSessionsValidFromDateTime
            RefreshTokensValidFromDateTime  = $u.AdditionalProperties['refreshTokensValidFromDateTime']
            # --- On-premises sync ---
            OnPremisesSyncEnabled           = $u.OnPremisesSyncEnabled
            OnPremisesImmutableId           = $u.OnPremisesImmutableId
            OnPremisesDomainName            = $u.OnPremisesDomainName
            OnPremisesSamAccountName        = $u.OnPremisesSamAccountName
            OnPremisesUserPrincipalName     = $u.OnPremisesUserPrincipalName
            OnPremisesLastSyncDateTime      = $u.OnPremisesLastSyncDateTime
            OnPremisesDistinguishedName     = $u.OnPremisesDistinguishedName
            # --- Locale ---
            UsageLocation                   = $u.UsageLocation
            PreferredLanguage               = $u.PreferredLanguage
            PreferredDataLocation           = $u.PreferredDataLocation
            # --- Sign-in activity ---
            LastSignInDateTime              = $lastSignIn
            LastNonInteractiveSignInDateTime = $lastNonInteractive
            # --- Summary counts (detail lives in the _* nested properties) ---
            GroupMembershipCount            = $groups.Count
            DirectoryRoleCount              = $dirRoles.Count
            AzureRbacAssignmentCount        = $azureRbac.Count
            AppRoleAssignmentCount          = @($apps.AppRoleAssignments).Count
            OAuth2GrantCount                = @($apps.OAuth2Grants).Count
            AuthMethodCount                 = $authMethods.Count
            LicenseCount                    = $licenses.Count
            DeviceCount                     = $devices.Count
            # --- Full nested data (for JSON / separate CSVs) ---
            _Groups                         = $groups
            _DirectoryRoles                 = $dirRoles
            _AzureRbac                      = $azureRbac
            _AppRoleAssignments             = @($apps.AppRoleAssignments)
            _OAuth2Grants                   = @($apps.OAuth2Grants)
            _AuthMethods                    = $authMethods
            _Licenses                       = $licenses
            _Devices                        = $devices
            _Manager                        = $orgInfo.Manager
            _DirectReports                  = $orgInfo.DirectReports
            _ExtensionAttributes            = $profileData.ExtensionAttrs
        }
    } catch {
        $Script:Stats.Errors++
        Write-Log "  [Export-SingleUser] ERROR for ${UserId}: $($_.Exception.Message)" 'ERROR'
        return $null
    }
}

# ---------------------------------------------------------------------------
# FILE EXPORT
# ---------------------------------------------------------------------------

function Write-JsonExport {
    param([Parameter(Mandatory)][object[]] $Users)

    $jsonDir = Join-Path $Script:OutputDir 'per_user'
    $null = New-Item -ItemType Directory -Force -Path $jsonDir

    # Full combined file
    $fullPath = Join-Path $Script:OutputDir 'users_full.json'
    $Users | ConvertTo-Json -Depth 30 | Set-Content -Path $fullPath -Encoding UTF8
    Write-Log '  users_full.json' 'SUCCESS'

    # Individual per-user files
    foreach ($u in $Users) {
        $safe = ($u.UserPrincipalName -replace '[\\/:*?"<>|@]','_')
        $u | ConvertTo-Json -Depth 30 |
             Set-Content -Path (Join-Path $jsonDir "$safe.json") -Encoding UTF8
    }
    Write-Log "  per_user/*.json  ($($Users.Count) files)" 'SUCCESS'
}

function Write-CsvExports {
    param([Parameter(Mandatory)][object[]] $Users)

    # Flat profile CSV (columns that don't start with '_')
    $profileCols = $Users[0].PSObject.Properties.Name |
                   Where-Object { -not $_.StartsWith('_') }
    $Users | Select-Object $profileCols |
        Export-Csv -Path (Join-Path $Script:OutputDir 'users_profile.csv') `
                   -NoTypeInformation -Encoding UTF8
    Write-Log '  users_profile.csv' 'SUCCESS'

    # Group memberships
    $rows = foreach ($u in $Users) {
        foreach ($g in ($u._Groups | Where-Object { $_['ObjectType'] -eq 'Group' })) {
            [PSCustomObject][ordered]@{
                UserPrincipalName        = $u.UserPrincipalName
                UserDisplayName          = $u.DisplayName
                GroupId                  = $g['Id']
                GroupDisplayName         = $g['DisplayName']
                Category                 = $g['Category']
                SecurityEnabled          = $g['SecurityEnabled']
                MailEnabled              = $g['MailEnabled']
                IsDynamic                = $g['IsDynamic']
                MembershipRule           = $g['MembershipRule']
                OnPremisesSynced         = $g['OnPremisesSynced']
                OnPremisesNetBiosName    = $g['OnPremisesNetBiosName']
                OnPremisesDomainName     = $g['OnPremisesDomainName']
                OnPremisesSamAccountName = $g['OnPremisesSamAccountName']
                OnPremisesDN             = $g['OnPremisesDN']
            }
        }
    }
    $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_groups.csv') `
                       -NoTypeInformation -Encoding UTF8
    Write-Log '  users_groups.csv' 'SUCCESS'

    # Entra directory roles
    $rows = foreach ($u in $Users) {
        foreach ($r in $u._DirectoryRoles) {
            [PSCustomObject][ordered]@{
                UserPrincipalName = $u.UserPrincipalName
                UserDisplayName   = $u.DisplayName
                RoleDisplayName   = $r['RoleDisplayName']
                RoleDescription   = $r['RoleDescription']
                RoleDefinitionId  = $r['RoleDefinitionId']
                DirectoryScopeId  = $r['DirectoryScopeId']
                AssignmentType    = $r['AssignmentType']
                Source            = $r['Source']
            }
        }
    }
    $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_directory_roles.csv') `
                       -NoTypeInformation -Encoding UTF8
    Write-Log '  users_directory_roles.csv' 'SUCCESS'

    # Azure RBAC (optional)
    if ($Script:IncludeAzureRBAC) {
        $rows = foreach ($u in $Users) {
            foreach ($r in $u._AzureRbac) {
                [PSCustomObject][ordered]@{
                    UserPrincipalName = $u.UserPrincipalName
                    UserDisplayName   = $u.DisplayName
                    RoleName          = $r['RoleDefinitionName']
                    RoleDefinitionId  = $r['RoleDefinitionId']
                    Scope             = $r['Scope']
                    ResourceGroupName = $r['ResourceGroupName']
                    ResourceName      = $r['ResourceName']
                    ResourceType      = $r['ResourceType']
                    CanDelegate       = $r['CanDelegate']
                    AssignmentId      = $r['AssignmentId']
                }
            }
        }
        $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_azure_rbac.csv') `
                           -NoTypeInformation -Encoding UTF8
        Write-Log '  users_azure_rbac.csv' 'SUCCESS'
    }

    # Enterprise app role assignments
    $rows = foreach ($u in $Users) {
        foreach ($a in $u._AppRoleAssignments) {
            [PSCustomObject][ordered]@{
                UserPrincipalName  = $u.UserPrincipalName
                UserDisplayName    = $u.DisplayName
                ApplicationName    = $a['ApplicationName']
                ServicePrincipalId = $a['ServicePrincipalId']
                AppRoleName        = $a['AppRoleName']
                AppRoleId          = $a['AppRoleId']
                CreatedDateTime    = $a['CreatedDateTime']
                AssignmentId       = $a['AssignmentId']
            }
        }
    }
    $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_app_assignments.csv') `
                       -NoTypeInformation -Encoding UTF8
    Write-Log '  users_app_assignments.csv' 'SUCCESS'

    # OAuth2 grants
    $rows = foreach ($u in $Users) {
        foreach ($g in $u._OAuth2Grants) {
            [PSCustomObject][ordered]@{
                UserPrincipalName = $u.UserPrincipalName
                UserDisplayName   = $u.DisplayName
                ClientAppName     = $g['ClientAppName']
                ClientId          = $g['ClientId']
                ResourceId        = $g['ResourceId']
                Scope             = $g['Scope']
                ConsentType       = $g['ConsentType']
            }
        }
    }
    $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_oauth2_grants.csv') `
                       -NoTypeInformation -Encoding UTF8
    Write-Log '  users_oauth2_grants.csv' 'SUCCESS'

    # MFA / auth methods
    if (-not $Script:SkipAuthMethods) {
        $rows = foreach ($u in $Users) {
            foreach ($m in $u._AuthMethods) {
                [PSCustomObject][ordered]@{
                    UserPrincipalName = $u.UserPrincipalName
                    UserDisplayName   = $u.DisplayName
                    MethodType        = $m['MethodType']
                    PhoneNumber       = $m['PhoneNumber']
                    EmailAddress      = $m['EmailAddress']
                    DisplayName       = $m['DisplayName']
                    AaGuid            = $m['AaGuid']
                    CreatedDateTime   = $m['CreatedDateTime']
                }
            }
        }
        $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_auth_methods.csv') `
                           -NoTypeInformation -Encoding UTF8
        Write-Log '  users_auth_methods.csv' 'SUCCESS'
    }

    # Licenses
    $rows = foreach ($u in $Users) {
        foreach ($l in $u._Licenses) {
            [PSCustomObject][ordered]@{
                UserPrincipalName = $u.UserPrincipalName
                UserDisplayName   = $u.DisplayName
                SkuPartNumber     = $l['SkuPartNumber']
                SkuId             = $l['SkuId']
            }
        }
    }
    $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_licenses.csv') `
                       -NoTypeInformation -Encoding UTF8
    Write-Log '  users_licenses.csv' 'SUCCESS'

    # Devices
    if (-not $Script:SkipDevices) {
        $rows = foreach ($u in $Users) {
            foreach ($d in $u._Devices) {
                [PSCustomObject][ordered]@{
                    UserPrincipalName      = $u.UserPrincipalName
                    UserDisplayName        = $u.DisplayName
                    Relationship           = $d['Relationship']
                    DeviceId               = $d['DeviceId']
                    DeviceDisplayName      = $d['DisplayName']
                    OperatingSystem        = $d['OperatingSystem']
                    OsVersion              = $d['OsVersion']
                    TrustType              = $d['TrustType']
                    IsCompliant            = $d['IsCompliant']
                    IsManaged              = $d['IsManaged']
                    ProfileType            = $d['ProfileType']
                    EnrollmentType         = $d['EnrollmentType']
                    ManagementType         = $d['ManagementType']
                    RegisteredDateTime     = $d['RegisteredDateTime']
                    ApproximateLastSignIn  = $d['ApproximateLastSignIn']
                    OnPremisesSynced       = $d['OnPremisesSynced']
                }
            }
        }
        $rows | Export-Csv -Path (Join-Path $Script:OutputDir 'users_devices.csv') `
                           -NoTypeInformation -Encoding UTF8
        Write-Log '  users_devices.csv' 'SUCCESS'
    }
}

function Write-SummaryReport {
    param([Parameter(Mandatory)][object[]] $Users)

    $elapsed    = (Get-Date) - $Script:StartTime
    $elapsedFmt = '{0:D2}h {1:D2}m {2:D2}s' -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds

    # Pre-compute group stats outside the here-string to avoid inline script blocks
    $groupTotals = $Users | ForEach-Object { $_.GroupMembershipCount } |
                   Measure-Object -Sum -Average -Maximum
    $groupSum  = $groupTotals.Sum
    $groupAvg  = '{0:N1}' -f $groupTotals.Average
    $groupMax  = $groupTotals.Maximum

    $rbacIncluded    = if ($Script:IncludeAzureRBAC)     { '[+]' } else { '[-]' }
    $authIncluded    = if (-not $Script:SkipAuthMethods) { '[+]' } else { '[-]' }
    $devIncluded     = if (-not $Script:SkipDevices)     { '[+]' } else { '[-]' }

    $rbacFileLine    = if ($Script:IncludeAzureRBAC) {
        '  users_azure_rbac.csv          - Azure resource RBAC role assignments'
    } else {
        '  (Azure RBAC not collected - run with -IncludeAzureRBAC to enable)'
    }
    $authFileLine    = if (-not $Script:SkipAuthMethods) {
        '  users_auth_methods.csv        - MFA and authentication methods'
    } else {
        '  (Auth methods not collected - remove -SkipAuthMethods to enable)'
    }
    $devFileLine     = if (-not $Script:SkipDevices) {
        '  users_devices.csv             - Registered and joined device details'
    } else {
        '  (Devices not collected - remove -SkipDevices to enable)'
    }

    $scopeValueLine = if ($Script:ScopeValue) { $Script:ScopeValue } else { '(entire tenant)' }

    $report = @"
================================================================================
  Entra ID User Export - Summary Report
  Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  Duration   : $elapsedFmt
  Script v   : $($Script:Version)
  Tenant     : $($Script:TenantName)  [$($Script:TenantId)]
================================================================================

  SCOPE
  -------------------------------------------------------------------------------
  Type            : $($Script:ScopeType)
  Value           : $scopeValueLine
  Include Guests  : $($Script:IncludeGuests)

  STATISTICS
  -------------------------------------------------------------------------------
  Users found     : $($Script:Stats.TotalUsers)
  Users exported  : $($Script:Stats.UsersProcessed)
  Errors          : $($Script:Stats.Errors)
  Skipped         : $($Script:Stats.SkippedUsers)

  GROUP MEMBERSHIP SUMMARY
  -------------------------------------------------------------------------------
  Total group memberships  : $groupSum
  Average per user         : $groupAvg
  Most groups (one user)   : $groupMax

  DATA COLLECTED PER USER
  -------------------------------------------------------------------------------
  [+] Full user profile (50+ attributes incl. on-prem sync, ImmutableId, DN)
  [+] Transitive group memberships (Security / M365 / Distribution / AU)
  [+] Entra ID directory role assignments (direct + PIM)
  $rbacIncluded Azure Resource RBAC across all accessible subscriptions
  [+] Enterprise app role assignments
  [+] OAuth2 delegated permission grants (user consents)
  $authIncluded MFA and authentication methods
  [+] Assigned licenses by SKU
  $devIncluded Registered and Entra-joined devices
  [+] Manager and direct reports
  [+] Sign-in activity timestamps (interactive + non-interactive)
  [+] Extension / custom attributes

  OUTPUT FILES  ->  $($Script:OutputDir)
  -------------------------------------------------------------------------------
  users_full.json               - complete nested data (all users combined)
  per_user/<upn>.json           - one JSON file per user (nested)
  users_profile.csv             - flat user profile attributes
  users_groups.csv              - group memberships (incl. on-prem AD info)
  users_directory_roles.csv     - Entra ID directory role assignments
$rbacFileLine
  users_app_assignments.csv     - enterprise application role assignments
  users_oauth2_grants.csv       - OAuth2 delegated permission grants
$authFileLine
  users_licenses.csv            - license assignments by SKU
$devFileLine
  export_summary.txt            - this summary report
  export.log                    - detailed run-time log
================================================================================
"@

    $reportPath = Join-Path $Script:OutputDir 'export_summary.txt'
    $report | Set-Content -Path $reportPath -Encoding UTF8
    Write-Host $report -ForegroundColor White
    Write-Log "Summary report saved: $reportPath" 'SUCCESS'
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

function Main {
    Write-Banner

    # Determine scope and options (interactive when -Scope not supplied)
    if (-not $PSBoundParameters.ContainsKey('Scope')) {
        $sel = Show-ScopeMenu
        $Script:ScopeType  = $sel.Scope
        $Script:ScopeValue = if ($sel.Value) { $sel.Value } else { '' }
        Show-OptionsMenu
    } else {
        $Script:ScopeType        = $Scope
        $Script:ScopeValue       = if ($ScopeValue) { $ScopeValue } else { '' }
    }

    # Add auth-methods scope now that options are resolved
    if (-not $Script:SkipAuthMethods) {
        $Script:GraphScopes += 'UserAuthenticationMethod.Read.All'
    }

    # Create output directory and start log
    $null = New-Item -ItemType Directory -Force -Path $Script:OutputDir
    $Script:LogFile = Join-Path $Script:OutputDir 'export.log'
    Write-Log "Output directory: $($Script:OutputDir)"

    # Install / load modules
    Install-RequiredModules -Modules $Script:RequiredGraphModules -Label 'Microsoft Graph modules'
    if ($Script:IncludeAzureRBAC) {
        Install-RequiredModules -Modules $Script:RequiredAzModules -Label 'Azure PowerShell modules'
    }

    # Connect
    Connect-ToEntra
    # After a successful Entra connection the script-level TenantId is populated;
    # pass it through to Azure RM so both sessions target the same tenant.
    if ($Script:IncludeAzureRBAC) { Connect-ToAzure }

    # Resolve user IDs for the selected scope
    Write-Log 'Resolving user list ...'
    $userIds = Get-UserIdsByScope -ScopeType $Script:ScopeType -ScopeVal $Script:ScopeValue

    if (-not $userIds -or $userIds.Count -eq 0) {
        Write-Log 'No users found for the selected scope. Exiting.' 'WARN'
        return
    }
    $Script:Stats.TotalUsers = $userIds.Count
    Write-Log "Found $($userIds.Count) user(s) to export." 'SUCCESS'

    # Collect data per user
    Write-Log 'Starting per-user data collection ...'
    $allUsers = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($uid in $userIds) {
        $i++
        $result = Export-SingleUser -UserId $uid -Index $i -Total $userIds.Count
        if ($null -ne $result) { $allUsers.Add($result) }
    }
    Write-Progress -Activity 'Exporting Entra Users' -Completed

    if ($allUsers.Count -eq 0) {
        Write-Log 'No user data collected - nothing to export.' 'WARN'
        return
    }

    $arr = $allUsers.ToArray()

    # Write output files
    Write-Log "Writing output to: $($Script:OutputDir) ..."
    if ($Script:ExportFormat -in 'JSON','Both') {
        Write-Log 'Writing JSON files ...'
        Write-JsonExport -Users $arr
    }
    if ($Script:ExportFormat -in 'CSV','Both') {
        Write-Log 'Writing CSV files ...'
        Write-CsvExports -Users $arr
    }

    Write-SummaryReport -Users $arr
    Write-Log "Export complete! All files saved to: $($Script:OutputDir)" 'SUCCESS'
}

# Entry point
Main
