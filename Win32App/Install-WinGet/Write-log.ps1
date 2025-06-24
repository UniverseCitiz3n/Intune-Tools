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
    } else {
        $callerNameCommand = ''
    }
    if ($callerFunction.ScriptName) {
        $callerNameScript = Split-Path -Leaf $callerFunction.ScriptName
    } else {
        $callerNameScript = ''
    }
    if ($callerFunction.ScriptLineNumber) {
        $callerLineNumber = $callerFunction.ScriptLineNumber
    } else {
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