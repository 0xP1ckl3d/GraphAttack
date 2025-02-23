 _______________________________________________________________________________
                                                                               
        ____                 _        _   _   _             _                  
       / ___|_ __ __ _ _ __ | |__    / \ | |_| |_ __ _  ___| | __              
      | |  _| '__/ _` | '_ \| '_ \  / _ \| __| __/ _` |/ __| |/ /              
      | |_| | | | (_| | |_) | | | |/ ___ \ |_| || (_| | (__|   <               
       \____|_|  \__,_| .__/|_| |_/_/   \_\__|\__\__,_|\___|_|\_\              
                      |_|                                                      
_______________________________________________________________________________

                       M I C R O S O F T   G R A P H
               E N U M E R A T I O N  &  A T T A C K  S C R I P T

--------------------------------------------------------------------------------
CONTENTS
--------------------------------------------------------------------------------
1.  Introduction
2.  Prerequisites
3.  Installation
4.  Script Overview
5.  Functions in Detail
    5.1  Connect-Graph
    5.2  Get-UpdatableGroups
    5.3  Add-SelfToGroup
    5.4  Remove-SelfFromGroup
    5.5  Get-SharePointSiteURLs
    5.6  Invoke-GraphRecon
    5.7  Get-SecurityGroups
    5.8  Invoke-DumpCAPS
    5.9  Invoke-DumpApps
    5.10 Get-DynamicGroups
    5.11 Get-AzureADUsers
    5.12 Invoke-InviteGuest
    5.13 Invoke-DriveFileDownload
    5.14 Invoke-SearchSharePointAndOneDrive
    5.15 Invoke-SearchUserAttributes
    5.16 Invoke-SearchMailbox
    5.17 Invoke-SearchTeamsMessages
    5.18 Invoke-GraphEnum
6.  Example Usage
7.  Troubleshooting & Common Issues
8.  Licence

--------------------------------------------------------------------------------
1. INTRODUCTION
--------------------------------------------------------------------------------
This PowerShell script is based on https://github.com/dafthack/GraphRunner. 
The main difference is the ability to authenticate via interactive auth. This 
script is a rewrite of Graphrunner for Microsoft Graph PowerShell.
The script handles majority of the same functionality as GraphRunner:

• Retrieving and manipulating group membership
• Searching SharePoint, OneDrive, and mailboxes
• Enumerating conditional access policies, app registrations, and more
• Performing organisational and user reconnaissance

_Note, that testing of all features has been limited, so feel free to
log an issue for bugs or feature requests._

Optimised for Microsoft Graph PowerShell v2.25.0.

--------------------------------------------------------------------------------
2. PREREQUISITES
--------------------------------------------------------------------------------
• A Windows or cross-platform PowerShell environment (PowerShell 5.1+ or 7.x).
• Permissions to install modules if not already present.
• An account with sufficient Azure AD / Microsoft 365 permissions (e.g. 
  Global Reader, Security Reader, or relevant delegated permissions).
• Internet access to reach Microsoft Graph endpoints.

--------------------------------------------------------------------------------
3. INSTALLATION
--------------------------------------------------------------------------------
1) Save this .ps1 script locally.

2) Load the script in PowerShell:

   `Import-Module .\graphattack.ps1`

3) Once loaded, you can call the defined functions directly in the same 
   PowerShell session.

--------------------------------------------------------------------------------
4. SCRIPT OVERVIEW
--------------------------------------------------------------------------------
• Installs missing Microsoft Graph modules automatically.
• Imports them for usage in the current session.
• Provides multiple distinct functions for enumerating and auditing:
  - AD Groups (including dynamic membership checks)
  - Security groups
  - SharePoint & OneDrive search
  - Guest invitations
  - App registrations & Enterprise apps
  - Mailbox and Teams message searching
• Does not execute anything automatically aside from module installations 
  and a basic connection check.

--------------------------------------------------------------------------------
5. FUNCTIONS IN DETAIL
--------------------------------------------------------------------------------

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.1  Connect-Graph
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    `Connect-Graph`

Description:
    Installs & imports required Microsoft Graph submodules if missing, then
    prompts you to sign in interactively (if not already authenticated).
    This is often the first function you’ll run.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.2  Get-UpdatableGroups
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Get-UpdatableGroups -Output "<YourOutputFile.csv>"

Description:
    Lists all groups in the tenant and checks whether you are allowed to update
    each group’s membership. Requires 'Group.Read.All' or similar permissions.

Parameter:
    -Output  (String)
       CSV file path for exporting details about which groups are "updatable."

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.3  Add-SelfToGroup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Add-SelfToGroup -GroupId <String> -Email <String>

Description:
    Adds the user to the specified group, given the group’s object ID and 
    users email.

Parameters:
    -GroupId  (String)
    -Email    (String)

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.4  Remove-SelfFromGroup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Remove-SelfFromGroup -GroupId <String> -Email <String>

Description:
    Removes the user from the specified group. Similar permission requirements 
    to Add-SelfToGroup.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.5  Get-SharePointSiteURLs
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Get-SharePointSiteURLs [-Output <String>]

Description:
    Queries SharePoint & OneDrive drives using the Graph Search API, returning
    the webUrl for each discovered site. If -Output is provided, the results go
    to CSV.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.6  Invoke-GraphRecon
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-GraphRecon

Description:
    Performs an overall reconnaissance:
    1) Retrieves organisation details (tenant ID, domains, etc.).
    2) Retrieves current user details.
    3) Reads default user role permissions from the authorisation policy.
    4) Uses 'estimateAccess' to summarise which high-level directory actions
       you are allowed to perform.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.7  Get-SecurityGroups
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Get-SecurityGroups [-OutputFile <String>]

Description:
    Retrieves all security-enabled groups and enumerates their members. Exports
    results (by default) to "security_groups.csv" if -OutputFile is specified.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.8  Invoke-DumpCAPS
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-DumpCAPS [-ResolveGuids]

Description:
    Dumps all conditional access policies in the tenant, printing:
    • Display name
    • State
    • Included/Excluded users, apps, platforms
    • Grant/session controls
    
  _The optional -ResolveGuids switch is a placeholder for future
  resolution of GUID-based references._

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.9  Invoke-DumpApps
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-DumpApps

Description:
    Enumerates all App Registrations and Enterprise Apps (service principals),
    along with any assigned permissions or app role assignments. Requires
    'Application.Read.All' and 'Directory.Read.All'.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.10 Get-DynamicGroups
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Get-DynamicGroups [-OutputPath <String>]

Description:
    Finds all groups that have a dynamic membership rule, then uses
    estimateAccess to see whether you can update them. Exports grouped results
    (allowed / conditional / denied) to CSV.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.11 Get-AzureADUsers
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Get-AzureADUsers [-OutFile <String>]

Description:
    Retrieves all userPrincipalNames (UPNs) in your Azure AD tenant. Useful for
    a quick user enumeration. Exports them to a text file (e.g., "users.txt").

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.12 Invoke-InviteGuest
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-InviteGuest -DisplayName <String> -EmailAddress <String> 
                       [-RedirectUrl <String>] [-SendInvitationMessage <Bool>] 
                       [-CustomMessageBody <String>]

Description:
    Sends a guest user invitation email to an external address. By default, the
    user is taken to the MyApps portal to accept the invitation.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.13 Invoke-DriveFileDownload
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-DriveFileDownload -DriveItemIDs <String> -FileName <String> 
                             [-Tokens <Object[]>]

Description:
    Downloads a single file from a drive (OneDrive/SharePoint) using a combined
    driveId:itemId string. Used internally by the search function, but you can
    call it manually as well.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.14 Invoke-SearchSharePointAndOneDrive
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-SearchSharePointAndOneDrive -SearchTerm <String> [-ResultCount <Int>]
                                       [-UnlimitedResults] [-OutFile <String>]
                                       [-ReportOnly]

Description:
    Searches SharePoint & OneDrive for files matching a search term (including
    KQL operators like "password AND filetype:xlsx"). Optionally downloads
    matching files if you confirm. Exports results to CSV if -OutFile is given.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.15 Invoke-SearchUserAttributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-SearchUserAttributes -SearchTerm <String> [-OutFile <String>]

Description:
    Retrieves ALL users, enumerates various attributes (e.g., displayName, mail,
    jobTitle), and checks if the given search term appears. Exports matches 
    to CSV if requested.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.16 Invoke-SearchMailbox
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-SearchMailbox -SearchTerm <String> [-MessageCount <Int>] 
                         [-OutFile <String>] [-PageResults]

Description:
    Searches your mailbox for emails containing the specified term in subject,
    body, or other fields. Exports findings to CSV if -OutFile is provided, and
    can fetch multiple pages if -PageResults is set.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.17 Invoke-SearchTeamsMessages
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-SearchTeamsMessages -KeyPhrase <String> [-BatchSize <Int>] 
                               [-OutputFile <String>] [-FetchAll]

Description:
    Lists Teams channels the signed-in user can access, retrieving messages that
    contain the specified phrase. Can save them to CSV if -OutputFile is set
    and fetch all results with -FetchAll.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5.18 Invoke-GraphEnum
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Usage:
    Invoke-GraphEnum [-DetectorFile <String>] [-DisableRecon] [-DisableUsers] 
                     [-DisableGroups] [-DisableEmail] [-DisableSharePoint] 
                     [-DisableTeams] [-Delay <Int>] [-Jitter <Double>]

Description:
    A "master" function that performs a series of enumerations in one pass:
      1) Organisation & user recon
      2) User listing
      3) Security groups listing
      4) Email, SharePoint/OneDrive, Teams searching
    The -DetectorFile can contain custom search queries. You can skip
    components by specifying the respective -Disable switches.

_Note: The detector file implementation is not yet functional. This will be
fixed in future releases._

--------------------------------------------------------------------------------
6. EXAMPLE USAGE
--------------------------------------------------------------------------------
• Connect to Graph, then run reconnaissance:
  Connect-Graph
  Invoke-GraphRecon

• Export updatable groups to CSV:
  Get-UpdatableGroups -Output "Updatable_Groups.csv"

• Add or remove yourself from a group:
  Add-SelfToGroup -GroupId "00000000-aaaa-bbbb-cccc-111111111111" -Email "user@tenant"
  Remove-SelfFromGroup -GroupId "00000000-aaaa-bbbb-cccc-111111111111" -Email "user@tenant"

• Enumerate security groups, export to CSV:
  Get-SecurityGroups -OutputFile "SecurityGroups.csv"

• Full SharePoint & OneDrive search for “password”:
  Invoke-SearchSharePointAndOneDrive -SearchTerm "password" -UnlimitedResults

• Search your mailbox for “secret” and export results:
  Invoke-SearchMailbox -SearchTerm "secret" -OutFile "SecretEmails.csv" -PageResults

• Run the “master” enumeration (detectors in a JSON file):
  Invoke-GraphEnum -DetectorFile "detectors.json"

--------------------------------------------------------------------------------
7. TROUBLESHOOTING & COMMON ISSUES
--------------------------------------------------------------------------------
1) **Permissions**:
   If any function fails (e.g., “Access denied”), ensure you have the correct 
   roles or admin consents (Directory.Read.All, Group.Read.All, etc.). Permissions
   can be added by running (for example):

   ```Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"```

3) **Rate Limits**:
   For large tenants, you may encounter HTTP 429 (rate limit). The script may 
   pause and retry. If this persists, try smaller queries or run off-peak.

4) **Module Installation**:
   The script attempts to install missing modules. Check your environment’s
   policy if installations fail.

--------------------------------------------------------------------------------
8. LICENCE
--------------------------------------------------------------------------------
This script is provided under the MIT Licence (unless otherwise indicated by your
organisation’s policies). You are free to modify and distribute it per licence
terms.

