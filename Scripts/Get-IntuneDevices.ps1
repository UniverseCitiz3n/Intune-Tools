<#
.SYNOPSIS
    Retrieves Intune-managed device information from a specific Azure AD group.

.DESCRIPTION
    This script retrieves detailed information about devices managed by Microsoft Intune (Microsoft Endpoint Manager). 
    It queries devices based on Azure AD group memberships using the Microsoft Graph API. It optionally supports authentication via an existing Microsoft Graph API access token.

.EXAMPLE
    .\Get-IntuneDevices.ps1

    Prompts for a Group ID or Azure AD URL and returns Intune-managed device details.
.NOTES
    Author: Maciej Horbacz
    Version: 1.0.0
    PowerShell Gallery: https://www.powershellgallery.com/packages/Get-IntuneDevices
#>
function Get-IntuneDevices {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [string]$AccessToken
    )


    # Establish connection
    if ($AccessToken) {
        $AccessToken = $AccessToken.Trim()
        $headers = @{
            "Authorization" = "$AccessToken"
            "Content-Type"  = "application/json"
        }
    }

    # Get group ID from input
    $groupInput = Read-Host "Enter the Entra ID group link (containing 'groupid/XXXXXXXXXXXX') or a group GUID"
    if ($groupInput -match "groupid/([0-9a-fA-F\-]+)") {
        $groupId = $matches[1]
        Write-Host "Detected group ID: $groupId"
    } elseif ($groupInput -match "^(?:\{)?[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}(?:\})?$") {
        $groupId = $groupInput
        Write-Host "Detected group GUID: $groupId"
    } else {
        Write-Error "Invalid group link or GUID."
        return
    }

    # Retrieve group members
    try {
        if ($AccessToken) {
            $groupMembers = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/$groupId/members" -Headers $headers -Method GET).value
        } else {
            $groupMembers = Get-MgGroupMember -GroupId $groupId -All
        }
    } catch {
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
            }
            try {
                if ($AccessToken) {
                    Write-Host "Processing device ID: $deviceId"
                    # Use the custom filter instead of device-specific filter
                    $filter = "contains(azureADDeviceId, '$deviceId')" 
                    $uri = "https://graph.microsoft.com/beta/deviceManagement/manageddevices?`$filter=$filter"
                    $deviceDetail = (Invoke-RestMethod -Uri $uri -Headers $headers -Method GET).value
                } else {
                    $deviceDetail = Get-MgDeviceManagementManagedDevice -Filter "deviceId eq '$deviceId'"
                }
            } catch {
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
        } elseif ($memberType -eq "#microsoft.graph.user") {
            $userUpn = $member.userPrincipalName
            try {
                if ($AccessToken) {
                    $deviceDetail = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=UserPrincipalName eq '$userUpn'" -Headers $headers -Method GET).value
                } else {
                    $deviceDetail = (Get-MgDeviceManagementManagedDevice -Filter "UserPrincipalName eq '$userUpn'").value
                }
            } catch {
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
        } else {
            Write-Host "Skipping member with id $($member.id) and type $memberType."
        }
    }

    Write-Host "Retrieved $($deviceResults.Count) device record(s)."
    return $deviceResults
}

# Usage example:
# $devices = Get-IntuneDevices
# $devices | Format-Table -AutoSize

# Initialize global variables
$global:devices = $null
$global:AccessToken = $null
Clear-Host
do {
    Write-Host ""
    Write-Host "Select an option:"
    Write-Host "1. Connect to Tenant (or use AccessToken)"
    Write-Host "2. Disconnect from Tenant"
    Write-Host "3. Check group"
    Write-Host "4. Show devices as table"
    Write-Host "5. Show devices as list"
    Write-Host "6. Export devices to CSV"
    Write-Host "7. Exit"
    $choice = Read-Host "Enter choice (1-7)"
    
    switch ($choice) {
        "1" {
            $inputValue = Read-Host "Enter your tenant ID, domain, or paste AccessToken"
            if ($inputValue -match "\.") {
                # The input likely is an AccessToken (contains dot separator)
                $global:AccessToken = $inputValue.Trim()
                Write-Host "AccessToken stored."
            } else {
                # Assume the input is a tenant ID or domain
                try {
                    Connect-MgGraph -TenantId $inputValue -Scopes "DeviceManagementManagedDevices.Read.All", "Group.Read.All", "User.Read.All", "GroupMember.Read.All" -ErrorAction Stop
                    $global:AccessToken = $null
                } catch {
                    Write-Error "Failed to connect: $_"
                }
            }
        }
        "2" {
            try {
                Disconnect-MgGraph
                Write-Host "Disconnected from tenant."
                $global:AccessToken = $null
            } catch {
                Write-Error "Failed to disconnect: $_"
            }
        }
        "3" {
            $global:devices = Get-IntuneDevices -AccessToken $global:AccessToken
            if ($global:devices) {
                Write-Host "Device details retrieved."
            }
        }
        "4" {
            if (@($global:devices).Count -gt 0) {
                $global:devices | Format-Table -AutoSize
            } else {
                Write-Error "No device data available. Run option 3 first."
            }
        }
        "5" {
            if (@($global:devices).Count -gt 0) {
                $global:devices | Format-List *
            } else {
                Write-Error "No device data available. Run option 3 first."
            }
        }
        "6" {
            if ($global:devices -and $global:devices.Count -gt 0) {
                $csvPath = Read-Host "Enter the full path for CSV export"
                try {
                    $global:devices | Export-Csv -Path $csvPath -NoTypeInformation -Force
                    Write-Host "Exported device details to $csvPath."
                } catch {
                    Write-Error "Export failed: $_"
                }
            } else {
                Write-Error "No device data available. Run option 3 first."
            }
        }
        "7" {
            $response = Read-Host "Do you want to keep devices in current session? (y/n)"
            if ($response -notmatch '^(y|Y)$') {
                $global:devices = $null
            } else {
                Write-Host "Device details are in '`$devices'"
            }
            $global:AccessToken = $null
            exit
        }
        default {
            Write-Host "Invalid option. Try again." 
        }
    }
} while ($true)
