$Global:ErrorActionPreference = 'Stop'
# Parameters

$FailCode = 1
$SuccessCode = 0

# Log - Set location based on execution context
$System = (new-object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)).Translate([System.Security.Principal.NTAccount])
$logLocation = if ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -eq $System.Value) {
	"C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
}
else {
	"C:\ProgramData\XXXX"
}
$Filelog = "$logLocation\ApplicationInstall.log"
$component = "Winget Installation"

#Info
. $PSScriptRoot\Write-log.ps1

#Custom exit
function Exit-WithCode {
	param
	(
		$exitcode
	)

	$host.SetShouldExit($exitcode)
}

If (-not(Test-Path -Path $logLocation -PathType Container)) {
	New-Item -Path $logLocation -ItemType Directory
	Write-log -logLocation $Filelog -Message "Log location folder created" -component $component -Type Info
}
#######################################################################
#Install
#Write-log -logLocation $Filelog -Message "Installation with arguments: $ArgumentListInstal" -component $component -Type Info
Try {
	Write-log -logLocation $Filelog -Message "Installing UI xaml" -component $component -Type Info
	Add-AppxProvisionedPackage -Online -PackagePath "$PSScriptRoot\Microsoft.UI.Xaml.2.8_8.2501.31001.0_x64.appx" -SkipLicense | Out-Null 
	Write-log -logLocation $Filelog -Message "Installing VCLibs" -component $component -Type Info
	Add-AppxProvisionedPackage -Online -PackagePath "$PSScriptRoot\Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64.appx" -SkipLicense | Out-Null 
	Write-log -logLocation $Filelog -Message "Installing Desktop App Installer" -component $component -Type Info
	Add-AppxProvisionedPackage -Online -PackagePath "$PSScriptRoot\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -SkipLicense | Out-Null 
	New-Item -Path $logLocation -Name "Winget.success" -ItemType File -Force | Out-Null
	Write-log -logLocation $Filelog -Message "Exit with code $SuccessCode" -component $component -Type Info
	Exit-WithCode -exitcode $SuccessCode
}
Catch {
	$Err = [PSCustomObject]@{
		Exception = $_.Exception.Message
		Reason    = $_.CategoryInfo.Reason
		Target    = $_.CategoryInfo.TargetName
		Script    = $_.InvocationInfo.ScriptName
		Line      = $_.InvocationInfo.ScriptLineNumber
		Column    = $_.InvocationInfo.OffsetInLine
	}
	Write-log -logLocation $Filelog -Message "Script ERROR" -component $component -Type Error
	Write-log -logLocation $Filelog -Message "$Err" -component $component -Type Error
	Write-log -logLocation $Filelog -Message "Exit with code $FailCode" -component $component -Type Error
	Exit-WithCode -exitcode $FailCode
}