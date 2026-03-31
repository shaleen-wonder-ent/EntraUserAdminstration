# Entra ID Comprehensive User Export Tool

A PowerShell script that performs a **complete, migration-ready export** of Microsoft Entra ID (formerly Azure AD) users. Every attribute and relationship needed to recreate identities in another Identity Provider (IdP) is collected and written to structured JSON and CSV files.

---

## Table of Contents

- [Overview](#overview)
- [What Data Is Collected](#what-data-is-collected)
- [Export Scopes](#export-scopes)
- [Prerequisites](#prerequisites)
- [Required Permissions](#required-permissions)
- [Installation](#installation)
- [Usage](#usage)
  - [Interactive Mode](#interactive-mode)
  - [Command-Line Mode](#command-line-mode)
  - [Parameters](#parameters)
  - [Examples](#examples)
- [Output Files](#output-files)
- [Sample Output Structure](#sample-output-structure)
- [Troubleshooting](#troubleshooting)
- [Security Notes](#security-notes)

---

## Overview

`Export-EntraUsers.ps1` connects to Microsoft Graph (and optionally Azure Resource Manager) to export a complete snapshot of user identities from Entra ID. You choose the **scope** — an entire tenant, a security group, a department, an administrative unit, or individual users — and the script collects every piece of identity data and writes it to a timestamped output folder.

Use cases:
- **IdP migration** (migrating from Entra ID to Okta, Ping, Auth0, AWS IAM Identity Center, etc.)
- **Tenant-to-tenant migration**
- **Compliance auditing** — RBAC and group membership snapshots
- **User onboarding automation** — seed a new system with rich user profiles

---

## What Data Is Collected

| Category | Details |
|---|---|
| **Full user profile** | 50+ attributes: UPN, display name, given name, surname, mail, proxy addresses, phone numbers, job title, department, company, employee ID/type, office location |
| **Address** | Street, city, state, country, postal code |
| **Account state** | `accountEnabled`, `userType`, `externalUserState`, created date, password last change date, password policies |
| **Sign-in activity** | Last interactive sign-in, last non-interactive sign-in |
| **On-premises sync** | `onPremisesImmutableId`, `onPremisesSamAccountName`, `onPremisesDistinguishedName`, domain name, last sync timestamp, provisioning errors |
| **Group memberships** | Transitive — Security groups, Microsoft 365 groups, Distribution lists, Administrative Units — includes AD on-prem metadata and dynamic membership rules |
| **Entra directory roles** | All assigned roles (Global Admin, Exchange Admin, etc.) via unified RBAC + legacy directory role objects |
| **Azure Resource RBAC** | Role assignments across all accessible subscriptions (optional) — role name, scope, resource group, resource name |
| **Enterprise app assignments** | Which enterprise applications the user has access to and their role within each app |
| **OAuth2 delegated grants** | User-consented delegated permissions per application |
| **MFA / Auth methods** | Microsoft Authenticator, FIDO2 security keys, Phone (SMS/voice), Email OTP, Software OATH, Windows Hello for Business, Temporary Access Pass, Password |
| **Licenses** | Assigned SKUs (E3, E5, Business Premium, etc.) with per-service-plan provisioning status |
| **Devices** | Registered (BYOD) and Entra-joined devices — OS, version, compliance state, management type, last sign-in, Intune enrollment type |
| **Manager & direct reports** | Manager UPN + all direct reports |
| **Custom / extension attributes** | Directory schema extensions and open extensions |

---

## Export Scopes

| Scope | Description |
|---|---|
| **Tenant** | All member users in the directory |
| **SecurityGroup** | Transitive members of a specific security group (name or Object ID) |
| **M365Group** | Transitive members of a Microsoft 365 group (name or Object ID) |
| **Department** | Users filtered by the `Department` attribute (exact match) |
| **AdministrativeUnit** | Members of an Administrative Unit (name or Object ID) |
| **SpecificUsers** | Comma-separated list of UPNs or Object IDs |

---

## Prerequisites

### PowerShell Version

- **PowerShell 5.1** or later (Windows PowerShell or PowerShell 7+)
- To check: `$PSVersionTable.PSVersion`

### PowerShell Modules

The script **auto-installs** any missing modules from the PowerShell Gallery using `Install-Module -Scope CurrentUser`. No manual installation needed on first run.

Modules used:

| Module | Purpose |
|---|---|
| `Microsoft.Graph.Authentication` | Graph authentication |
| `Microsoft.Graph.Users` | User profile data |
| `Microsoft.Graph.Groups` | Group memberships |
| `Microsoft.Graph.Identity.DirectoryManagement` | Administrative Units, directory roles |
| `Microsoft.Graph.Identity.Governance` | PIM role assignments |
| `Microsoft.Graph.Applications` | App role assignments, OAuth2 grants |
| `Microsoft.Graph.Identity.SignIns` | Authentication methods |
| `Az.Accounts` *(optional)* | Azure RM authentication (only if `-IncludeAzureRBAC`) |
| `Az.Resources` *(optional)* | Azure RBAC role assignment queries (only if `-IncludeAzureRBAC`) |

---

## Required Permissions

The account used to run the script (or the app registration if using app-only auth) must have:

### Microsoft Graph

| Permission | Required? | Purpose |
|---|---|---|
| `User.Read.All` | Required | Read all user profiles |
| `Group.Read.All` | Required | Read group properties |
| `GroupMember.Read.All` | Required | Read group memberships |
| `Directory.Read.All` | Required | Read directory objects (AUs, roles) |
| `RoleManagement.Read.All` | Required | Read RBAC role assignments |
| `RoleManagement.Read.Directory` | Required | Read directory-scoped role assignments |
| `AuditLog.Read.All` | Required | Read sign-in activity |
| `Application.Read.All` | Required | Read app registrations and service principals |
| `UserAuthenticationMethod.Read.All` | Optional | Read MFA/auth methods (omit `-SkipAuthMethods` to use) |

> **Note:** These permissions will be requested automatically during the interactive sign-in prompt. A Global Reader or User Administrator role covers most of these in practice.

### Azure Resource Manager (optional, `-IncludeAzureRBAC` only)

- **Reader** role on each subscription you want to scan for RBAC assignments

---

## Installation

```powershell
# Option 1 - Clone the repository
git clone https://github.com/shaleen-wonder-ent/EntraUserAdminstration.git
Set-Location EntraUserAdminstration

# Option 2 - Download the script directly
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/shaleen-wonder-ent/EntraUserAdminstration/main/Export-EntraUsers.ps1" `
                  -OutFile "Export-EntraUsers.ps1"
```

If your execution policy blocks unsigned scripts, run once from an elevated PowerShell session:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Usage

### Interactive Mode

Run the script with no parameters. You will be shown menus to:
1. Choose the export scope
2. Configure options (guests, Azure RBAC, auth methods, devices, format)
3. Sign in to Microsoft Graph (browser pop-up)

```powershell
.\Export-EntraUsers.ps1
```

### Command-Line Mode

Pass parameters directly for unattended or scheduled runs.

```powershell
.\Export-EntraUsers.ps1 -Scope Tenant -ExportFormat Both
```

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Scope` | String | *(interactive)* | Export scope: `Tenant`, `SecurityGroup`, `M365Group`, `Department`, `AdministrativeUnit`, `SpecificUsers` |
| `-ScopeValue` | String | | Name or Object ID of the target group/AU/department. Comma-separated UPNs for `SpecificUsers`. || `-TenantId` | String | | Your Entra ID Tenant ID (GUID) or primary domain (e.g. `contoso.onmicrosoft.com`). **Required if you have a personal @outlook.com / @hotmail.com account on the same browser profile.** Scopes the sign-in prompt to your work/school tenant and prevents the MSA consumer error. Find it at [Entra admin center](https://entra.microsoft.com) → Overview → Tenant ID. || `-OutputDir` | String | `.\EntraExport_<timestamp>` | Folder where all output files are written |
| `-ExportFormat` | String | `Both` | `JSON`, `CSV`, or `Both` |
| `-IncludeAzureRBAC` | Switch | `$false` | Collect Azure Resource RBAC assignments across all subscriptions |
| `-SkipAuthMethods` | Switch | `$false` | Skip MFA / authentication method collection |
| `-SkipDevices` | Switch | `$false` | Skip device registration collection |
| `-IncludeGuests` | Switch | `$false` | Include Guest users (default exports Members only) |

### Examples

```powershell
# 1. Interactive mode (recommended for first run)
.\Export-EntraUsers.ps1

# 1a. If you have a personal Microsoft account on the same browser profile,
#     supply -TenantId to force sign-in against your work tenant and avoid
#     the "not supported for MSA accounts" error.
.\Export-EntraUsers.ps1 -TenantId "contoso.onmicrosoft.com"
.\Export-EntraUsers.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# 2. Entire tenant, all formats + Azure RBAC
.\Export-EntraUsers.ps1 -Scope Tenant -ExportFormat Both -IncludeAzureRBAC

# 3. All members of a security group (by name)
.\Export-EntraUsers.ps1 -Scope SecurityGroup -ScopeValue "SG-Finance-All"

# 4. All members of a security group (by Object ID)
.\Export-EntraUsers.ps1 -Scope SecurityGroup -ScopeValue "a1b2c3d4-1234-5678-abcd-ef1234567890"

# 5. All members of a Microsoft 365 group
.\Export-EntraUsers.ps1 -Scope M365Group -ScopeValue "All Company"

# 6. All users in the Engineering department
.\Export-EntraUsers.ps1 -Scope Department -ScopeValue "Engineering"

# 7. Members of an Administrative Unit
.\Export-EntraUsers.ps1 -Scope AdministrativeUnit -ScopeValue "AU-EMEA-Users"

# 8. Two specific users, JSON only
.\Export-EntraUsers.ps1 -Scope SpecificUsers `
    -ScopeValue "alice@contoso.com,bob@contoso.com" `
    -ExportFormat JSON

# 9. Entire tenant including guests, custom output folder
.\Export-EntraUsers.ps1 -Scope Tenant -IncludeGuests `
    -OutputDir "C:\Exports\EntraMigration_2026"

# 10. Tenant export skipping devices and auth methods (faster for large tenants)
.\Export-EntraUsers.ps1 -Scope Tenant -SkipDevices -SkipAuthMethods -ExportFormat CSV
```

---

## Output Screenshot

After a successful run the output folder will look like this in VS Code (all files marked **U** — untracked, kept out of git by `.gitignore`):

![Export output folder in VS Code](docs/output-screenshot.png)

The timestamped folder (`EntraExport_<yyyyMMdd_HHmmss>/`) contains:

| File | Description |
|---|---|
| `users_profile.csv` | Flat user profile — one row per user |
| `users_groups.csv` | Group memberships |
| `users_directory_roles.csv` | Entra directory role assignments |
| `users_app_assignments.csv` | Enterprise app role assignments |
| `users_oauth2_grants.csv` | Delegated OAuth2 permission grants |
| `users_auth_methods.csv` | MFA / authentication methods |
| `users_licenses.csv` | License SKU assignments |
| `users_devices.csv` | Registered / joined devices |
| `export_summary.txt` | Run statistics and file manifest |
| `export.log` | Detailed timestamped log |

---

## Output Files

All files are written to the output directory (default: `.\EntraExport_<yyyyMMdd_HHmmss>\`).

```
EntraExport_20260331_142500/
|
|-- users_full.json             <- All users, fully nested (one JSON array)
|
|-- per_user/                   <- One JSON file per user
|   |-- alice_contoso.com.json
|   |-- bob_contoso.com.json
|   `-- ...
|
|-- users_profile.csv           <- Flat user profile (50+ columns, one row per user)
|-- users_groups.csv            <- Group memberships (one row per user+group pair)
|-- users_directory_roles.csv   <- Entra ID directory role assignments
|-- users_azure_rbac.csv        <- Azure resource RBAC (if -IncludeAzureRBAC)
|-- users_app_assignments.csv   <- Enterprise app role assignments
|-- users_oauth2_grants.csv     <- Delegated OAuth2 permission grants
|-- users_auth_methods.csv      <- MFA and auth methods (if not -SkipAuthMethods)
|-- users_licenses.csv          <- License SKU assignments
|-- users_devices.csv           <- Registered/joined devices (if not -SkipDevices)
|
|-- export_summary.txt          <- Run statistics and file manifest
`-- export.log                  <- Detailed timestamped log
```

### File Details

#### `users_profile.csv`
One row per user. Key columns:

| Column | Description |
|---|---|
| `ObjectId` | Entra ID object ID (GUID) |
| `UserPrincipalName` | UPN — the primary login identifier |
| `OnPremisesImmutableId` | ImmutableId used to match cloud user to on-prem AD account |
| `OnPremisesDistinguishedName` | Full AD Distinguished Name |
| `OnPremisesSamAccountName` | Legacy SAM account name |
| `LastSignInDateTime` | Last interactive sign-in (from audit log) |
| `PasswordPolicies` | e.g. `DisablePasswordExpiration` |
| `AccountEnabled` | `True` / `False` |

#### `users_groups.csv`
One row per user-group membership pair.

| Column | Description |
|---|---|
| `UserPrincipalName` | User identifier |
| `GroupDisplayName` | Group name |
| `Category` | `SecurityGroup`, `M365Group`, `DistributionList`, `Other` |
| `IsDynamic` | Whether the group uses a dynamic membership rule |
| `MembershipRule` | The dynamic membership rule expression (if applicable) |
| `OnPremisesDN` | AD Distinguished Name of the group (if synced from on-prem) |

#### `users_directory_roles.csv`
One row per user-role assignment.

| Column | Description |
|---|---|
| `RoleDisplayName` | e.g. `Global Administrator`, `Exchange Administrator` |
| `DirectoryScopeId` | `/` for tenant-wide, or resource-scoped ID for AU-scoped roles |
| `AssignmentType` | `Assigned` |
| `Source` | `UnifiedRBAC` (PIM/direct) or `DirectoryRole` (legacy) |

#### `users_azure_rbac.csv`
One row per Azure resource RBAC assignment (only when `-IncludeAzureRBAC` is used).

| Column | Description |
|---|---|
| `RoleName` | e.g. `Owner`, `Contributor`, `Reader` |
| `Scope` | Full ARM scope path |
| `ResourceGroupName` | Resource group (if assignment is resource-group scoped) |
| `ResourceName` | Individual resource name (if assignment is resource-scoped) |

---

## Sample Output Structure

### `users_full.json` (truncated)

```json
[
  {
    "ObjectId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "UserPrincipalName": "alice@contoso.com",
    "DisplayName": "Alice Smith",
    "Department": "Engineering",
    "AccountEnabled": true,
    "OnPremisesImmutableId": "ABC123==",
    "OnPremisesSamAccountName": "asmith",
    "LastSignInDateTime": "2026-03-28T09:14:22Z",
    "GroupMembershipCount": 7,
    "DirectoryRoleCount": 0,
    "LicenseCount": 1,
    "_Groups": [
      {
        "DisplayName": "SG-Engineering-All",
        "Category": "SecurityGroup",
        "OnPremisesDomainName": "contoso.local",
        "OnPremisesSamAccountName": "SG-Engineering-All"
      }
    ],
    "_AuthMethods": [
      { "MethodType": "MicrosoftAuthenticator", "DisplayName": "Alice's iPhone" },
      { "MethodType": "Phone", "PhoneNumber": "+1 555-0100", "PhoneType": "mobile" }
    ],
    "_Licenses": [
      { "SkuPartNumber": "SPE_E3" }
    ]
  }
]
```

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| **`This API is not supported for MSA accounts`** | You signed in with a personal Microsoft account (@outlook.com / @hotmail.com) instead of a work/school Entra ID account | Re-run with `-TenantId`: `.\Export-EntraUsers.ps1 -TenantId "contoso.onmicrosoft.com"`. Your Tenant ID is in [Entra admin center](https://entra.microsoft.com) → Overview → Tenant ID |
| `Install-Module` fails | No internet access or restricted policy | Pre-install modules manually: `Install-Module Microsoft.Graph -Scope CurrentUser` |
| Sign-in window does not appear | Running in a non-interactive session | Run from an interactive PowerShell session, or use a service principal with a client certificate |
| `Insufficient privileges` error | Missing Graph permissions | Grant the listed permissions in Entra ID -> App registrations or use a Global Reader account |
| `User not found` warnings in log | User deleted between listing and collection | Non-fatal; user is skipped and counted in `Errors` in the summary |
| `[AuthMethods] Error` for every user | Missing `UserAuthenticationMethod.Read.All` | Either grant the permission or add `-SkipAuthMethods` |
| Azure RBAC collection very slow | Large number of subscriptions | Expected; consider scoping to specific subscriptions using `Set-AzContext` before running |
| Output CSV has empty columns | Some attributes not populated in the tenant | Normal — not all attributes are used by every organisation |
| `Export-Csv` writes no rows for a table | No users had that type of data (e.g. no app assignments) | Normal — empty CSVs are still created as placeholders |

---

## Security Notes

- **Credentials are never written to disk.** The script uses Microsoft Graph's interactive or device-code authentication flow — tokens are stored in memory only for the duration of the run.
- **Output files may contain sensitive PII** (phone numbers, addresses, sign-in timestamps). Store the output folder in a secure location and apply appropriate access controls before sharing.
- **Least-privilege recommendation:** Use a dedicated service account with the minimum Graph permissions listed above rather than a Global Administrator account.
- The script performs **read-only** operations only — no users, groups, or roles are modified.

---

## Version History

| Version | Date | Changes |
|---|---|---|
| 2.0.0 | 2026-03-31 | Initial public release — full IdP migration export |

---

## License

MIT License — see [LICENSE](LICENSE) for details.
