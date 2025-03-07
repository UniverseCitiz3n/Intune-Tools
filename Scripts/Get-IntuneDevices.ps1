<#PSScriptInfo
.VERSION
1.0.4

.GUID
e00cc407-4231-4af7-a226-f2a9b28395f3

.AUTHOR
Maciej Horbacz

.SYNOPSIS
Retrieves Intune device details via the Microsoft Graph API.


.DESCRIPTION
This script leverages the Microsoft Graph API to retrieve detailed information about devices managed by Intune. It allows users to specify an Entra ID group (group can contain users and/or devices), from which it extracts device details, which can be displayed as a table or list, or exported to a CSV file.


.COMPANYNAME
Cloud Aligned

.COPYRIGHT
(c) 2025 Maciej. All rights reserved.

.TAGS
Intune, MicrosoftGraph, Devices, EntraID

.LICENSEURI

.PROJECTURI
https://github.com/UniverseCitiz3n/Intune-Tools

.ICONURI

.EXTERNALMODULEDEPENDENCIES 
Microsoft.Graph

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
    v1.0.0 - Initial version.
    v1.0.1 - Minor bug fix.
    v1.0.2 - Minor bug fix.
    v1.0.3 - Minor bug fix.
    v1.0.4 - Microsoft.Graph as required.
   
.PRIVATEDATA
#>
function Get-IntuneDevices {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [string]$AccessToken,
        [Parameter(Mandatory = $false)]
        [string]$TenantID
    )

    # Ensure only one connection parameter is used
    if ($AccessToken -and $TenantID) {
        Write-Error "Provide either AccessToken or TenantID, not both."
        return
    }

    if ($AccessToken) {
        $AccessToken = $AccessToken.Trim()
        $headers = @{
            "Authorization" = "$AccessToken"
            "Content-Type"  = "application/json"
        }
    }
    elseif ($TenantID) {
        # No additional headers needed as the connection is already established using TenantID
    }
    else {
        Write-Error "No connection information provided. Run menu option 1 first."
        return
    }

    # Get group ID from input
    $groupInput = Read-Host "Enter the Entra ID group link (containing 'groupid/XXXXXXXXXXXX') or a group GUID"
    if ($groupInput -match "groupid/([0-9a-fA-F\-]+)") {
        $groupId = $matches[1]
        Write-Host "Detected group ID: $groupId"
    }
    elseif ($groupInput -match "^(?:\{)?[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}(?:\})?$") {
        $groupId = $groupInput
        Write-Host "Detected group GUID: $groupId"
    }
    else {
        Write-Error "Invalid group link or GUID."
        return
    }

    # Retrieve group members
    try {
        if ($AccessToken) {
            $groupMembers = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/$groupId/members" -Headers $headers -Method GET).value
        }
        else {
            # First, retrieve a temporary list to determine member types
            $tempMembers = Get-MgGroupMember -GroupId $groupId -All
            $hasUser = $false
            $hasDevice = $false
            foreach ($m in $tempMembers) {
                if ($m.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user") {
                    $hasUser = $true
                }
                elseif ($m.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.device") {
                    $hasDevice = $true
                }
            }

            # Retrieve members with appropriate functions based on detected types
            $groupMembers = @()
            if ($hasUser) {
                $users = Get-MgGroupMemberAsUser -GroupId $groupId -All
                foreach ($user in $users) {
                    if (-not $user.'@odata.type') {
                        $user | Add-Member -NotePropertyName '@odata.type' -NotePropertyValue "#microsoft.graph.user" -Force
                    }
                }
                $groupMembers += $users
            }
            if ($hasDevice) {
                $devices = Get-MgGroupMemberAsDevice -GroupId $groupId -All
                foreach ($device in $devices) {
                    if (-not $device.'@odata.type') {
                        $device | Add-Member -NotePropertyName '@odata.type' -NotePropertyValue "#microsoft.graph.device" -Force
                    }
                }
                $groupMembers += $devices
            }
        }
    }
    catch {
        Write-Error "Error retrieving group members: $_"
        return
    }

    $deviceResults = @()

    foreach ($member in $groupMembers) {
        $memberType = $member.'@odata.type'
        if (-not $memberType -and $member.userPrincipalName) {
            $memberType = "#microsoft.graph.user"
        }
        if ($memberType -eq "#microsoft.graph.device") {
            # Use 'deviceId' property from device object rather than 'id'
            $deviceId = $member.deviceId
            if (-not $deviceId) {
                Write-Error "Device ID not found for member $($member.id). Skipping." 
                continue
            }
            try {
                if ($AccessToken) {
                    Write-Host "Processing device ID: $deviceId"
                    # Use the custom filter instead of device-specific filter
                    $filter = "contains(azureADDeviceId, '$deviceId')" 
                    $uri = "https://graph.microsoft.com/beta/deviceManagement/manageddevices?`$filter=$filter"
                    $deviceDetail = (Invoke-RestMethod -Uri $uri -Headers $headers -Method GET).value
                }
                else {
                    Write-Host "Processing device ID: $deviceId"
                    $filter = "contains(azureADDeviceId, '$deviceId')"
                    $deviceDetail = Get-MgDeviceManagementManagedDevice -Filter $filter
                }
            }
            catch {
                Write-Warning "Unable to get Intune details for device id $deviceId"
                continue
            }
            foreach ($dev in $deviceDetail) {
                $deviceResults += [PSCustomObject]@{
                    DeviceName     = $dev.deviceName
                    Ownership      = $dev.managedDeviceOwnerType
                    Compliance     = $dev.complianceState
                    OS             = $dev.operatingSystem
                    OSVersion      = $dev.osVersion
                    PrimaryUserUPN = $dev.userPrincipalName
                    LastCheckIn    = $dev.lastSyncDateTime
                    EnrollmentDate = $dev.enrolledDateTime
                    Model          = $dev.model
                    Manufacturer   = $dev.manufacturer
                    SerialNumber   = $dev.serialNumber
                    JoinType       = $dev.joinType
                    EntraID        = $dev.azureADDeviceId
                }
            }
        }
        elseif ($memberType -eq "#microsoft.graph.user") {
            $userUpn = $member.userPrincipalName
            try {
                if ($AccessToken) {
                    Write-Host "Processing user: $userUpn"
                    $deviceDetail = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=UserPrincipalName eq '$userUpn'" -Headers $headers -Method GET).value
                }
                else {
                    Write-Host "Processing user: $userUpn"
                    $deviceDetail = Get-MgDeviceManagementManagedDevice -Filter "UserPrincipalName eq '$userUpn'"
                }
            }
            catch {
                Write-Warning "Unable to get devices for user $userUpn"
                continue
            }
            foreach ($dev in $deviceDetail) {
                $deviceResults += [PSCustomObject]@{
                    DeviceName     = $dev.deviceName
                    Ownership      = $dev.managedDeviceOwnerType
                    Compliance     = $dev.complianceState
                    OS             = $dev.operatingSystem
                    OSVersion      = $dev.osVersion
                    PrimaryUserUPN = $dev.userPrincipalName
                    LastCheckIn    = $dev.lastSyncDateTime
                    EnrollmentDate = $dev.enrolledDateTime
                    Model          = $dev.model
                    Manufacturer   = $dev.manufacturer
                    SerialNumber   = $dev.serialNumber
                    JoinType       = $dev.joinType
                    EntraID        = $dev.azureADDeviceId
                }
            }
        }
        else {
            Write-Host "Skipping member with id $($member.id) and type $memberType."
        }
    }

    Write-Host "Retrieved $($deviceResults.Count) device record(s)."
    return $deviceResults
}

# Initialize variables
$devices = $null
$AccessToken = $null
$Tenant = $null
Clear-Host
do {
    Write-Host ""
    Write-Host "Select an option:"
    Write-Host "1. Set Tenant or AccessToken"
    Write-Host "2. Disconnect from Tenant"
    Write-Host "3. Check group"
    Write-Host "4. Show devices as table"
    Write-Host "5. Show devices as list"
    Write-Host "6. Export devices to CSV"
    Write-Host "7. Exit"
    $choice = Read-Host "Enter choice (1-7)"
    
    switch ($choice) {
        "1" {
            $inputValue = Read-Host "Enter your AccessToken (if it contains a dot) or your tenant ID/domain"
            if ($inputValue -match "\.") {
                $AccessToken = $inputValue.Trim()
                $Tenant = $null
                Write-Host "AccessToken stored."
            }
            else {
                try {
                    $AccessToken = $null
                    $Tenant = $inputValue.Trim()
                    Import-Module Microsoft.Graph.Authentication
                    Connect-MgGraph -NoWelcome -TenantId $Tenant -Scopes "DeviceManagementManagedDevices.Read.All", "Group.Read.All", "User.Read.All", "GroupMember.Read.All" -ErrorAction Stop
                    Write-Host "Connected to tenant $Tenant."
                }
                catch {
                    Write-Error "Failed to connect to Microsoft Graph: $_"
                    return
                }
            }
        }
        "2" {
            try {
                Disconnect-MgGraph
                Write-Host "Disconnected from tenant."
                $Tenant = $null
                $AccessToken = $null
            }
            catch {
                Write-Error "Failed to disconnect: $_"
            }
        }
        "3" {
            if (-not $AccessToken -and -not $Tenant) {
                Write-Host "No connection established. Please run option 1 first."
            }
            else {
                if ($AccessToken) {
                    $devices = Get-IntuneDevices -AccessToken $AccessToken
                }
                else {
                    $devices = Get-IntuneDevices -TenantID $Tenant
                }
                if ($devices) {
                    Write-Host "Device details retrieved."
                }
            }
        }
        "4" {
            if (-not $devices -or $devices.Count -eq 0) {
                Write-Host "No devices loaded. Please run option 3 first."
            }
            else {
                $devices | Format-Table -AutoSize
            }
        }
        "5" {
            if (-not $devices -or $devices.Count -eq 0) {
                Write-Host "No devices loaded. Please run option 3 first."
            }
            else {
                $devices | Format-List *
            }
        }
        "6" {
            if (-not $devices -or $devices.Count -eq 0) {
                Write-Host "No devices loaded. Please run option 3 first."
            }
            else {
                $csvPath = Read-Host "Enter the full path for CSV export"
                try {
                    $devices | Export-Csv -Path $csvPath -NoTypeInformation -Force
                    Write-Host "Exported device details to $csvPath."
                }
                catch {
                    Write-Error "Export failed: $_"
                }
            }
        }
        "7" {
            $response = Read-Host "Do you want to keep devices in current session? (y/n)"
            if ($response -notmatch '^(y|Y)$') {
                $devices = $null
            }
            else {
                Write-Host 'Device details are in $devices'
            }
            $AccessToken = $null
            exit
        }
        default {
            Write-Host "Invalid option. Try again." 
        }
    }
} while ($true)