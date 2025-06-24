function Write-log {

    [CmdletBinding()]
    Param(
        [parameter(Mandatory = $false)]
        [String]$logLocation,

        [parameter(Mandatory = $true)]
        [String]$Message,

        [parameter(Mandatory = $false)]
        [String]$component = $component,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error")]
        [String]$Type
    )

    switch ($Type) {
        "Info" {
            [int]$Type = 1 
        }
        "Warning" {
            [int]$Type = 2 
        }
        "Error" {
            [int]$Type = 3 
        }
    }

    # Create a log entry
    $callerFunction = (Get-PSCallStack)[1]

    if ($callerFunction.InvocationInfo -and $callerFunction.InvocationInfo.MyCommand -and $callerFunction.InvocationInfo.MyCommand.Name) {
        $callerNameCommand = $callerFunction.InvocationInfo.MyCommand.Name
    }
    else {
        $callerNameCommand = ''
    }
    if ($callerFunction.ScriptName) {
        $callerNameScript = Split-Path -Leaf $callerFunction.ScriptName
    }
    else {
        $callerNameScript = ''
    }
    if ($callerFunction.ScriptLineNumber) {
        $callerLineNumber = $callerFunction.ScriptLineNumber
    }
    else {
        $callerLineNumber = ''
    }

    if ($callerNameCommand -eq '' -and $callerNameScript -ne '') {
        $callerNameCommand = $callerNameScript
    }

    $Content = "<![LOG[[Line:$callerLineNumber] $Message]LOG]!>" + `
        "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " + `
        "date=`"$(Get-Date -Format "M-d-yyyy")`" " + `
        "component=`"$component`" " + `
        "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + `
        "type=`"$Type`" " + `
        "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " + `
        "source=`"$callerNameCommand`" " + `
        "callerLine=`"$callerLineNumber`" " + `
        "file=`"$callerNameScript`">"

    # Write the line to the log file
    Add-Content -Path $logLocation -Value $Content
}
#Parameters
$adtSession = @{
    # MS Graph variables
    ClientID     = 'XXXXXX-XXXX-XXXX-XXXXXXXXXXXX'
    TenantID     = 'XXXXXX-XXXX-XXXX-XXXXXXXXXXXX'
    ClientSecret = 'XXXXXX-XXXX-XXXX-XXXXXXXXXXXX'
}
$System = (new-object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)).Translate([System.Security.Principal.NTAccount])
$logLocation = if ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -eq $System.Value) {
    "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
}
else {
    "C:\ProgramData\XXXX"
}
$Filelog = "$logLocation\ApplicationDetection.log"
$component = "Application Detection"
If (-not(Test-Path -Path $logLocation -PathType Container)) {
	New-Item -Path $logLocation -ItemType Directory
	Write-log -logLocation $Filelog -Message "Log location folder created" -component $component -Type Info
}
#Get location of winget.exe file
$WingetInfo = (Get-Item "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_8wekyb3d8bbwe\winget.exe").VersionInfo | Sort-Object -Property FileVersionRaw
#If multiple versions, pick most recent one
$WinGetEXE = $WingetInfo[-1].FileName

try {
    Import-Module -Name Microsoft.Graph.Authentication
    $SecureSecret = ConvertTo-SecureString -String $adtSession.ClientSecret -AsPlainText -Force
    $Credential = new-object -typename System.Management.Automation.PSCredential -ArgumentList $adtSession.ClientID, $SecureSecret
    Write-log -logLocation $Filelog -Message "Connecting to Microsoft Graph" -component $component -Type Info
    Connect-MgGraph -TenantId $adtSession.TenantID -ClientSecretCredential $Credential -NoWelcome
    $AppID = (Split-Path $PSCommandPath -Leaf) -replace '_\d+\.ps1$', ''
    Write-log -logLocation $Filelog -Message "AppID: $AppID" -component $component -Type Info
    $AppDetails = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$AppID"
    Write-log -logLocation $Filelog -Message "App Version: $($AppDetails.displayVersion)" -component $component -Type Info   
    Write-log -logLocation $Filelog -Message "App Package ID: $($AppDetails.notes)" -component $component -Type Info
    $isInstalled = & $WinGetEXE 'list' "$($AppDetails.notes)" '--accept-source-agreements' '-s' 'winget' '-e'

    # Check if application was found
    if ($isInstalled[-1] -eq 'No installed package found matching input criteria.') {
        Write-log -logLocation $Filelog -Message 'Application was not detected' -component $component -Type Error
        Write-log -logLocation $Filelog -Message "Exiting with code 1" -component $component -Type Info
        Exit 1
    }
    
    # More robust parsing of winget output
    $packageFound = $false
    $installedVersion = $null
    
    # Look for the package in the output (skip header lines)
    foreach ($line in $isInstalled) {
        if ($line -match "^\s*\S+.*$($AppDetails.notes)") {
            $packageFound = $true
            # Use regex to extract version more reliably
            if ($line -match '\s+(\d+(?:\.\d+)*(?:\.\d+)*)\s*(?:\s+(\d+(?:\.\d+)*(?:\.\d+)*))?\s*$') {
                try {
                    $installedVersion = [version]$matches[1]
                    Write-log -logLocation $Filelog -Message "Found installed version: $installedVersion" -component $component -Type Info
                    break
                }
                catch {
                    Write-log -logLocation $Filelog -Message "Could not parse installed version from: $line" -component $component -Type Warning
                }
            }
        }
    }
    
    if (-not $packageFound) {
        Write-log -logLocation $Filelog -Message 'Application package not found in WinGet output' -component $component -Type Error
        Write-log -logLocation $Filelog -Message "Exiting with code 1" -component $component -Type Info
        Exit 1
    }
    
    if ($null -eq $installedVersion) {
        Write-log -logLocation $Filelog -Message 'Could not determine installed version - treating as detected' -component $component -Type Warning
        Write-log -logLocation $Filelog -Message "Exiting with code 0" -component $component -Type Info
        Write-Output 'Could not determine installed version - treating as detected'
        Exit 0
    }
    
    # Parse expected version from Intune
    try {
        $expectedVersion = [version]$AppDetails.displayVersion
    }
    catch {
        Write-log -logLocation $Filelog -Message "Could not parse expected version: $($AppDetails.displayVersion)" -component $component -Type Warning
        Write-log -logLocation $Filelog -Message "Exiting with code 0" -component $component -Type Info
        Write-Output 'Could not parse expected version - treating as detected'
        Exit 0
    }
    
    # Version comparison logic
    if ($expectedVersion -gt $installedVersion) {
        Write-log -logLocation $Filelog -Message "Application needs update: installed=$installedVersion, expected=$expectedVersion" -component $component -Type Info
        Write-log -logLocation $Filelog -Message "Exiting with code 1" -component $component -Type Info
        Exit 1  # Not detected (needs update)
    }
    elseif ($expectedVersion -eq $installedVersion) {
        Write-log -logLocation $Filelog -Message "Application is at expected version: $installedVersion" -component $component -Type Info
        Write-log -logLocation $Filelog -Message "Exiting with code 0" -component $component -Type Info
        Write-Output 'Detected and correct version' 
        Exit 0  # 
    }
    else {
        Write-log -logLocation $Filelog -Message "Application is newer than expected: installed=$installedVersion, expected=$expectedVersion" -component $component -Type Info
        Write-log -logLocation $Filelog -Message "Exiting with code 0" -component $component -Type Info
        Write-Output 'Detected and newer version'
        Exit 0  # Detected (newer version is acceptable)
    }
}
catch {
    $Err = [PSCustomObject] @{
        Exception = $PSItem.Exception.Message
        Reason    = $PSItem.CategoryInfo.Reason
        Script    = $PSItem.InvocationInfo.ScriptName
        Line      = $PSItem.InvocationInfo.ScriptLineNumber
        Column    = $PSItem.InvocationInfo.OffsetInLine
    }
    Write-log -logLocation $Filelog -Message "Error: $($Err.Exception)" -component $component -Type Error
    Write-log -logLocation $Filelog -Message "Exiting with code 1" -component $component -Type Info
    Exit 1
}