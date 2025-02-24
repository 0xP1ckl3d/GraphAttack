# Microsoft Graph Enumeration & Attack Script

 _______________________________________________________________________________
                                                                               
        ____                 _        _   _   _             _                  
       / ___|_ __ __ _ _ __ | |__    / \ | |_| |_ __ _  ___| | __              
      | |  _| '__/ _` | '_ \| '_ \  / _ \| __| __/ _` |/ __| |/ /              
      | |_| | | | (_| | |_) | | | |/ ___ \ |_| || (_| | (__|   <               
       \____|_|  \__,_| .__/|_| |_/_/   \_\__|\__\__,_|\___|_|\_\              
                      |_|                                                      
 ______________________________________________________________________________

                      M I C R O S O F T   G R A P H
             E N U M E R A T I O N  &  A T T A C K   S C R I P T

---

## Contents

1. [Introduction](#1-introduction)  
2. [Prerequisites](#2-prerequisites)  
3. [Installation](#3-installation)  
4. [Script Overview](#4-script-overview)  
5. [Functions in Detail](#5-functions-in-detail)  
   1. [Connection Functions](#51-connection-functions)  
   2. [Enumeration Functions](#52-enumeration-functions)  
   3. [Content Recon Functions](#53-content-recon-functions)  
   4. [Attack Functions](#54-attack-functions)  
6. [Example Usage](#6-example-usage)  
7. [Troubleshooting & Common Issues](#7-troubleshooting--common-issues)  
8. [Licence](#8-licence)

---

## 1. Introduction

This PowerShell script is inspired by [GraphRunner](https://github.com/dafthack/GraphRunner) but rewrites the functionality for the Microsoft Graph PowerShell modules with interactive authentication support. Key features include:

- Retrieving and manipulating group membership.
- Searching SharePoint, OneDrive, and mailboxes.
- Enumerating conditional access policies, app registrations, devices, tenant details, and more.
- Performing comprehensive organisational and user reconnaissance.

> **Note**: Testing of all features has been limited, so please log an issue for bugs or feature requests.  
> **Optimised for** Microsoft Graph PowerShell v2.25.0.

---

## 2. Prerequisites

- A Windows or cross-platform PowerShell environment (PowerShell 5.1+ or 7.x).  
- Permissions to install modules if not already present.  
- An account with sufficient Azure AD / Microsoft 365 permissions (e.g. Global Reader, Security Reader, or relevant delegated permissions).  
- Internet access to reach Microsoft Graph endpoints.

---

## 3. Installation

1. Save this `.ps1` script locally.
2. Load the script in PowerShell:

   ```powershell
   Import-Module .\graphattack.ps1
   ```

3. Once loaded, you can call the defined functions directly in the same PowerShell session.

Alternatively, you can run import this script from a remote source:

   ```powershell
   Invoke-RestMethod -Uri https://raw.githubusercontent.com/0xP1ckl3d/GraphAttack/refs/heads/main/GraphAttack.ps1 | Invoke-Expression
   ```

---

## 4. Script Overview

- Installs missing Microsoft Graph modules automatically (if you don’t already have them).
- Imports them for usage in the current session.
- Provides multiple distinct functions for enumerating and auditing:
  - Azure AD Groups (including dynamic membership checks)
  - Security groups and privileged roles
  - SharePoint & OneDrive search
  - Guest invitations and OAuth app injection
  - Devices, tenant settings, app registrations & enterprise apps
  - Mailbox and Teams message searching
- Offers both single-purpose functions and a “master” enumeration function.

> **Note**: The script does not execute anything automatically aside from module installations and a basic connection check.

---

## 5. Functions in Detail

### 5.1 Connection Functions

---

#### Connect-Graph

**Usage**:
```powershell
Connect-Graph
```

**Description**:  
- Installs and imports required Microsoft Graph submodules if missing, then prompts you to sign in interactively (if not already authenticated).
- This is often the first function you’ll run to ensure you’re ready to call other commands.

---

### 5.2 Enumeration Functions

---

#### Get-UpdatableGroups

**Usage**:
```powershell
Get-UpdatableGroups -Output "<YourOutputFile.csv>"
```

**Description**:  
- Lists all groups in the tenant and checks whether you (the signed-in account) are allowed to update each group’s membership.  
- Requires `Group.Read.All` or similar permissions.

**Parameter**:
- **-Output** (String):  
  CSV file path for exporting details about which groups are updatable.

---

#### Get-SecurityGroups

**Usage**:
```powershell
Get-SecurityGroups [-OutputFile <String>]
```

**Description**:  
- Retrieves all **security-enabled groups** and enumerates their members.  
- Exports results to `security_groups.csv` if `-OutputFile` is specified.

---

#### Get-DynamicGroups

**Usage**:
```powershell
Get-DynamicGroups [-OutputPath <String>]
```

**Description**:  
- Finds all groups that have a dynamic membership rule.  
- Uses “estimateAccess” to see whether you can update them.  
- Exports grouped results (allowed / conditional / denied) to CSV if you provide `-OutputPath`.

---

#### Get-AzureADUsers

**Usage**:
```powershell
Get-AzureADUsers [-OutFile <String>]
```

**Description**:  
- Retrieves all userPrincipalNames (UPNs) in your Azure AD tenant.  
- Useful for a quick user enumeration.  
- Exports them to a text file (e.g., `users.txt`) if `-OutFile` is specified.

---

#### Get-PrivilegedUsers

**Usage**:
```powershell
Get-PrivilegedUsers
```

**Description**:  
- Enumerates users or entities with privileged roles (e.g., Global Admin, Security Admin) in Azure AD.  
- Fetches all role assignments, resolves each principal (user, group, or service principal) and displays the assigned role.

---

#### Get-MFAStatus

**Usage**:
```powershell
Get-MFAStatus
```

**Description**:  
- Retrieves MFA (Multi-Factor Authentication) status for all users in Azure AD.  
- Checks if they have configured MFA methods (phone, authenticator, etc.).  
- **Note**: Requires `UserAuthenticationMethod.Read.All` for complete access.

---

#### Get-Devices

**Usage**:
```powershell
Get-Devices
```

**Description**:  
- Retrieves all registered devices in Azure AD, including assigned owners.  
- Exports the device list to `devices.csv` in the current directory by default.

---

#### Get-TenantEnumeration

**Usage**:
```powershell
Get-TenantEnumeration
```

**Description**:  
- Enumerates detailed Azure AD tenant settings, including:
  - Tenant ID, default domain, verified domains
  - Federation configuration for each domain
  - External collaboration settings
  - Security defaults
  - Licence assignments
- Useful for full tenant profiling.

---

#### Invoke-GraphRecon

**Usage**:
```powershell
Invoke-GraphRecon
```

**Description**:  
- Performs an overall reconnaissance in one go:
  1. Retrieves organisation details (tenant ID, domains, etc.).
  2. Retrieves the current user’s details.
  3. Reads the default user role permissions from the authorisation policy.
  4. Uses `estimateAccess` to summarise which high-level directory actions you are allowed to perform.

---

#### Invoke-DumpCAPS

**Usage**:
```powershell
Invoke-DumpCAPS [-ResolveGuids]
```

**Description**:  
- Dumps all conditional access policies in the tenant, printing:
  - Display name, state, included/excluded users/groups/apps/platforms
  - Grant/session controls
- `-ResolveGuids` is a placeholder for future expansions that resolve GUID-based references to names.

---

#### Invoke-DumpApps

**Usage**:
```powershell
Invoke-DumpApps
```

**Description**:  
- Enumerates all App Registrations and Enterprise Apps (service principals), along with any assigned permissions or app role assignments.  
- Requires `Application.Read.All` and `Directory.Read.All` roles.

---

#### Invoke-GraphEnum

**Usage**:
```powershell
Invoke-GraphEnum [-DetectorFile <String>] [-DisableRecon] [-DisableUsers] [-DisableGroups]
                 [-DisableEmail] [-DisableSharePoint] [-DisableTeams] [-Delay <Int>] [-Jitter <Double>]
```

**Description**:  
- A “master” function that performs a series of enumerations in one pass:
  1. Organisation & user recon
  2. User listing
  3. Security groups listing
  4. Email, SharePoint/OneDrive, Teams searching

- The `-DetectorFile` can contain custom search queries.
- You can skip components by specifying the respective `-Disable*` switches.
- **Note**: The detector file implementation is not yet functional, but planned for future releases.

---

### 5.3 Content Recon Functions

---

#### Get-SharePointSiteURLs

**Usage**:
```powershell
Get-SharePointSiteURLs [-Output <String>]
```

**Description**:  
- Queries SharePoint & OneDrive drives using the Graph Search API.
- Returns the `webUrl` for each discovered site.
- If `-Output` is provided, the results are exported to CSV.

---

#### Invoke-SearchSharePointAndOneDrive

**Usage**:
```powershell
Invoke-SearchSharePointAndOneDrive -SearchTerm <String> 
                                   [-ResultCount <Int>] 
                                   [-UnlimitedResults] 
                                   [-OutFile <String>] 
                                   [-ReportOnly]
```

**Description**:  
- Searches SharePoint & OneDrive for files matching a search term (including KQL operators like `"password AND filetype:xlsx"`).
- Optionally downloads matching files if you confirm.
- Exports results to CSV if `-OutFile` is given.
- `-UnlimitedResults` attempts to retrieve as many results as possible.

---

#### Invoke-DriveFileDownload

**Usage**:
```powershell
Invoke-DriveFileDownload -DriveItemIDs <String> -FileName <String>
                         [-Tokens <Object[]>]
```

**Description**:  
- Downloads a single file from a drive (OneDrive/SharePoint) using a combined `driveId:itemId` string.
- Used internally by the search function, but you can call it manually as well.

---

#### Invoke-SearchMailbox

**Usage**:
```powershell
Invoke-SearchMailbox -SearchTerm <String> 
                    [-MessageCount <Int>] 
                    [-OutFile <String>] 
                    [-PageResults]
```

**Description**:  
- Searches **your mailbox** for emails containing the specified term in subject, body, or other fields.
- Exports findings to CSV if `-OutFile` is provided.
- Can fetch multiple pages if `-PageResults` is set.

---

#### Invoke-SearchTeamsMessages

**Usage**:
```powershell
Invoke-SearchTeamsMessages -KeyPhrase <String> 
                           [-BatchSize <Int>] 
                           [-OutputFile <String>] 
                           [-FetchAll]
```

**Description**:  
- Lists Teams channels the signed-in user can access, retrieving messages that contain the specified phrase.
- Can export to CSV if `-OutputFile` is set and fetch all results with `-FetchAll`.

---

#### Invoke-SearchUserAttributes

**Usage**:
```powershell
Invoke-SearchUserAttributes -SearchTerm <String> [-OutFile <String>]
```

**Description**:  
- Retrieves **all users**, enumerates various attributes (`displayName`, `mail`, `jobTitle`, etc.), and checks if the given search term appears.
- Exports matches to CSV if requested.

---

### 5.4 Attack Functions

---

#### Add-SelfToGroup

**Usage**:
```powershell
Add-SelfToGroup -GroupId <String> -Email <String>
```

**Description**:  
- Adds **your user** to the specified group, given the group’s object ID and your email address.
- Permissions required: Typically `Group.ReadWrite.All`.

**Parameters**:
- **-GroupId** (String)
- **-Email** (String)

---

#### Remove-SelfFromGroup

**Usage**:
```powershell
Remove-SelfFromGroup -GroupId <String> -Email <String>
```

**Description**:  
- Removes **your user** from the specified group.
- Similar permission requirements to **Add-SelfToGroup**.

---

#### Invoke-InviteGuest

**Usage**:
```powershell
Invoke-InviteGuest -DisplayName <String> -EmailAddress <String> 
                  [-RedirectUrl <String>] 
                  [-SendInvitationMessage <Bool>] 
                  [-CustomMessageBody <String>]
```

**Description**:  
- Sends a guest user invitation email to an external address.
- By default, the user is taken to the MyApps portal to accept the invitation.

---

#### Invoke-InjectOAuthApp

**Usage**:
```powershell
Invoke-InjectOAuthApp -AppName <String> -ReplyUrl <String> 
                      [-Scope <String[]>] 
                      [-Tokens <Object[]>]
```

**Description**:  
- Automates the deployment of an App Registration in Azure AD.
- Creates an App Registration, assigns OAuth permissions, and generates a consent URL for you to use.
- Useful if portal access is restricted but you can still register apps via Graph.

**Parameters**:
- **-AppName** (String): The display name of the new App Registration.
- **-ReplyUrl** (String): The redirect URL where OAuth tokens will be sent.
- **-Scope** (String[]): Comma-separated Microsoft Graph permissions (e.g., `"Mail.Read","User.Read"`). If omitted, defaults to broad “backdoor” permissions.
- **-Tokens** (Object[]): If you have pre-authenticated tokens, you can supply them; otherwise, function attempts interactive login.

---

#### Invoke-SecurityGroupCloner

**Usage**:
```powershell
Invoke-SecurityGroupCloner
```

**Description**:  
- Clones a security group and copies its members to a newly created group.
- Optionally adds your current user (and any other specified user) to the cloned group.
- Useful for replicating membership quickly, or establishing a group with near-identical privileges.

---

## 6. Example Usage

Below are several quick examples of how to use this script:

1. **Connect to Microsoft Graph, then run high-level reconnaissance**:
   ```powershell
   Connect-Graph
   Invoke-GraphRecon
   ```

2. **Export updatable groups to CSV**:
   ```powershell
   Get-UpdatableGroups -Output "Updatable_Groups.csv"
   ```

3. **Add or remove yourself from a group**:
   ```powershell
   Add-SelfToGroup -GroupId "00000000-aaaa-bbbb-cccc-111111111111" -Email "user@tenant"
   Remove-SelfFromGroup -GroupId "00000000-aaaa-bbbb-cccc-111111111111" -Email "user@tenant"
   ```

4. **Enumerate security groups, export to CSV**:
   ```powershell
   Get-SecurityGroups -OutputFile "SecurityGroups.csv"
   ```

5. **Search SharePoint & OneDrive for “password”**:
   ```powershell
   Invoke-SearchSharePointAndOneDrive -SearchTerm "password" -UnlimitedResults
   ```

6. **Search your mailbox for “secret” and export results**:
   ```powershell
   Invoke-SearchMailbox -SearchTerm "secret" -OutFile "SecretEmails.csv" -PageResults
   ```

7. **Get a quick list of all Azure AD users**:
   ```powershell
   Get-AzureADUsers -OutFile "AllUsers.txt"
   ```

8. **Discover privileged user assignments**:
   ```powershell
   Get-PrivilegedUsers
   ```

9. **Check MFA status for all users**:
   ```powershell
   Get-MFAStatus
   ```

10. **Clone a security group, optionally add yourself**:
    ```powershell
    Invoke-SecurityGroupCloner
    ```

11. **Inject an OAuth app**:
    ```powershell
    Invoke-InjectOAuthApp -AppName "WinDefend365" -ReplyUrl "https://localhost/windefend" -Scope "User.Read","Mail.Read"
    ```

12. **Run the “master” enumeration**:
    ```powershell
    Invoke-GraphEnum -DetectorFile "detectors.json"
    ```

---

## 7. Troubleshooting & Common Issues

1. **Permissions**:  
   If any function fails (e.g., “Access denied”), ensure you have the correct roles or admin consent (e.g. `Directory.Read.All`, `Group.Read.All`). You can add permissions by running:
   ```powershell
   Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"
   ```

2. **Rate Limits**:  
   For large tenants, you may encounter HTTP 429 (rate limit) responses. The script attempts to pause and retry, but if it persists, try smaller queries or run off-peak.

3. **Module Installation**:  
   The script attempts to install missing modules. Check your environment’s policies if installations fail (e.g., `Set-PSRepository` or `Install-Module` constraints).

4. **Unsupported or Unrecognised Endpoints**:  
   Some older endpoints or newly introduced APIs may not be included in your installed Microsoft Graph module. Update your modules or check for alternative endpoints.

---

## 8. Licence

This script is provided under the MIT Licence (unless otherwise indicated by your organisation’s policies). You are free to modify and distribute it under these terms.
