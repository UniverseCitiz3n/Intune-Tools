<#
.SYNOPSIS
    Uploads a .intunewin package to Microsoft Intune.

.DESCRIPTION
    This script uploads a PSADTv4 packaged .intunewin file to Microsoft Intune
    using the Microsoft Graph API. It handles app creation, content upload,
    and all necessary Azure Storage operations.

.PARAMETER PackagePath
    The absolute path to the package folder containing the .intunewin file.

.PARAMETER BearerToken
    The Microsoft Graph API Bearer token with DeviceManagementApps.ReadWrite.All permission.

.PARAMETER InstallMode
    The deployment mode for installation. Valid values: Interactive, Silent, NonInteractive.
    Default: Interactive

.PARAMETER UninstallMode
    The deployment mode for uninstallation. Valid values: Interactive, Silent, NonInteractive.
    Default: Interactive

.PARAMETER Detection
    Optional path to a PowerShell detection script. If not provided, uses default detection.

.PARAMETER AppName
    The display name for the application in Intune.

.PARAMETER AppVersion
    The version string for the application.

.PARAMETER AppVendor
    The publisher/vendor name for the application.

.PARAMETER AppId
    The WinGet package ID (used in notes field).

.PARAMETER IconPath
    Optional path to a PNG icon file for the application.

.EXAMPLE
    .\Invoke-IntuneUpload.ps1 -PackagePath "C:\Apps\Install-Chrome-PSADTv4" -BearerToken "eyJ0..." -AppName "Chrome" -AppVersion "120.0" -AppVendor "Google" -AppId "Google.Chrome"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PackagePath,

    [Parameter(Mandatory = $true)]
    [string]$BearerToken,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [string]$InstallMode = 'Interactive',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [string]$UninstallMode = 'Interactive',

    [Parameter(Mandatory = $false)]
    [string]$Detection,

    [Parameter(Mandatory = $true)]
    [string]$AppName,

    [Parameter(Mandatory = $true)]
    [string]$AppVersion,

    [Parameter(Mandatory = $true)]
    [string]$AppVendor,

    [Parameter(Mandatory = $true)]
    [string]$AppId,

    [Parameter(Mandatory = $false)]
    [string]$IconPath,

    [Parameter(Mandatory = $false)]
    [string]$Description,

    [Parameter(Mandatory = $false)]
    [string]$UILanguage,

    [Parameter(Mandatory = $false)]
    [string]$InstallParams
)

#region Assembly References
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web
#endregion

#region Constants
$script:chunkSizeInBytes = 1024 * 1024 * 6  # 6MB chunks
$script:returnCodes = @(
    @{ "returnCode" = 0; "type" = "success" },
    @{ "returnCode" = 1707; "type" = "success" },
    @{ "returnCode" = 3010; "type" = "softReboot" },
    @{ "returnCode" = 1641; "type" = "hardReboot" },
    @{ "returnCode" = 1618; "type" = "retry" }
)
#endregion

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        'Info' {
            'White' 
        }
        'Warning' {
            'Yellow' 
        }
        'Error' {
            'Red' 
        }
        'Success' {
            'Green' 
        }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-IntuneWinDetectionXml {
    param([string]$IntuneWinPath)

    $zip = $null
    try {
        if (-not (Test-Path $IntuneWinPath)) {
            throw "IntuneWin file not found: $IntuneWinPath"
        }

        $directory = [System.IO.Path]::GetDirectoryName($IntuneWinPath)
        $zip = [IO.Compression.ZipFile]::OpenRead($IntuneWinPath)

        $detectionEntry = $zip.Entries | Where-Object { $_.Name -eq "detection.xml" }
        if ($detectionEntry) {
            $detectionXmlPath = Join-Path $directory "detection_temp.xml"
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($detectionEntry, $detectionXmlPath, $true)

            $zip.Dispose()
            $zip = $null

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            $detectionXml = [xml](Get-Content $detectionXmlPath)
            Remove-Item $detectionXmlPath -Force -ErrorAction SilentlyContinue

            Write-Log "Successfully extracted detection.xml metadata"
            return $detectionXml
        } else {
            throw "detection.xml not found in IntuneWin package"
        }
    } catch {
        Write-Log "Error extracting detection.xml: $($_.Exception.Message)" -Level "Error"
        return $null
    } finally {
        if ($zip) {
            try {
                $zip.Dispose() 
            } catch { 
            } 
        }
    }
}

function Get-IntuneWinContentFile {
    param(
        [string]$IntuneWinPath,
        [xml]$DetectionXml
    )

    $zip = $null
    try {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "IntuneUpload"
        if (-not (Test-Path -LiteralPath $tempRoot)) {
            $null = New-Item -ItemType Directory -Path $tempRoot -Force
        }
        $uniqueDirName = "Content_" + [System.Guid]::NewGuid().ToString("N").Substring(0, 8)
        $directory = Join-Path $tempRoot $uniqueDirName
        $null = New-Item -ItemType Directory -Path $directory -Force
        
        $script:tempExtractionDir = $directory
        $fileName = $DetectionXml.ApplicationInfo.FileName
        $contentFilePath = Join-Path $directory $fileName

        $maxRetries = 3
        $retryCount = 0

        while ($retryCount -lt $maxRetries) {
            try {
                if ($retryCount -gt 0) {
                    Start-Sleep -Milliseconds (1000 * $retryCount)
                    Write-Log "Retrying content file extraction (attempt $($retryCount + 1)/$maxRetries)..."
                }

                $zip = [IO.Compression.ZipFile]::OpenRead($IntuneWinPath)
                $contentEntry = $zip.Entries | Where-Object { $_.Name -eq $fileName }

                if ($contentEntry) {
                    if (Test-Path $contentFilePath) {
                        Remove-Item $contentFilePath -Force -ErrorAction SilentlyContinue
                    }

                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($contentEntry, $contentFilePath, $false)

                    $zip.Dispose()
                    $zip = $null

                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()

                    Write-Log "Extracted encrypted content file: $fileName"
                    return $contentFilePath
                } else {
                    throw "Content file '$fileName' not found in IntuneWin package"
                }
            } catch [System.IO.IOException] {
                $retryCount++
                if ($zip) {
                    try {
                        $zip.Dispose() 
                    } catch { 
                    } $zip = $null 
                }
                if ($retryCount -ge $maxRetries) {
                    throw "Failed to extract content file after $maxRetries attempts: $($_.Exception.Message)"
                }
                Write-Log "File access conflict, retrying..." -Level "Warning"
            }
        }
    } catch {
        Write-Log "Error extracting content file: $($_.Exception.Message)" -Level "Error"
        return $null
    } finally {
        if ($zip) {
            try {
                $zip.Dispose() 
            } catch { 
            } 
        }
    }
}

function Invoke-AzureStorageUpload {
    param(
        [string]$AzureStorageUri,
        [string]$FilePath,
        [string]$FileContainerUri
    )

    try {
        Write-Log "Starting chunked upload to Azure Storage..."

        $chunkSizeInBytes = 1024 * 1024 * 6  # 6MB chunks
        $fileSize = (Get-Item $FilePath).Length
        $chunkCount = [Math]::Ceiling($fileSize / $chunkSizeInBytes)
        $chunkIds = @()

        $binaryReader = New-Object System.IO.BinaryReader([System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite))
        $sasRenewalTimer = [System.Diagnostics.Stopwatch]::StartNew()

        for ($chunk = 0; $chunk -lt $chunkCount; $chunk++) {
            $chunkId = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($chunk.ToString("0000")))
            $chunkIds += $chunkId

            $start = $chunk * $chunkSizeInBytes
            $length = [Math]::Min($chunkSizeInBytes, $fileSize - $start)
            $bytes = $binaryReader.ReadBytes($length)

            $currentChunk = $chunk + 1
            Write-Log "Uploading chunk $currentChunk of $chunkCount"

            $uri = "$AzureStorageUri&comp=block&blockid=$chunkId"
            $iso = [Text.Encoding]::GetEncoding("iso-8859-1")
            $encodedBody = $iso.GetString($bytes)

            $headers = @{
                "x-ms-blob-type" = "BlockBlob"
                "Content-Type"   = "text/plain; charset=iso-8859-1"
            }

            $null = Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $encodedBody

            if ($currentChunk -lt $chunkCount -and $sasRenewalTimer.ElapsedMilliseconds -ge 450000) {
                Write-Log "Renewing SAS token..."
                $renewalUri = "$FileContainerUri/renewUpload"
                $updatedUri = Invoke-RestMethod -Uri $renewalUri -Method Post -Headers $headers -Body "{}"
                $AzureStorageUri = $updatedUri.azureStorageUri
                $sasRenewalTimer.Restart()
            }
        }

        Write-Log "Finalizing Azure Storage upload..."
        $uri = "$AzureStorageUri&comp=blocklist"
        $xml = '<?xml version="1.0" encoding="utf-8"?><BlockList>'
        foreach ($chunkId in $chunkIds) {
            $xml += "<Latest>$chunkId</Latest>"
        }
        $xml += '</BlockList>'

        $headers = @{ "Content-Type" = "text/xml" }
        $null = Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $xml

        Write-Log "Azure Storage upload completed successfully" -Level "Success"
        return $true
    } catch {
        Write-Log "Azure Storage upload failed: $($_.Exception.Message)" -Level "Error"
        return $false
    } finally {
        if ($binaryReader) {
            $binaryReader.Dispose() 
        }
    }
}

function Get-JwtTenantInfo {
    param([string]$JwtToken)

    try {
        $token = $JwtToken -replace '^Bearer\s+', ''
        if ([string]::IsNullOrEmpty($token)) {
            return $null 
        }

        $tokenParts = $token.Split('.')
        if ($tokenParts.Length -ne 3) {
            return $null 
        }

        $payload = $tokenParts[1]
        switch ($payload.Length % 4) {
            2 {
                $payload += '==' 
            }
            3 {
                $payload += '=' 
            }
        }

        $decodedBytes = [System.Convert]::FromBase64String($payload)
        $decodedJson = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
        $tokenData = ConvertFrom-Json $decodedJson

        $tenantInfo = @{
            TenantId          = $tokenData.tid
            TenantDomain      = $null
            UserPrincipalName = $tokenData.upn
            IsValid           = $false
        }

        if ($tokenData.tid) {
            $tenantInfo.IsValid = $true 
        }
        if ($tokenData.upn -match '@(.+)$') {
            $tenantInfo.TenantDomain = $matches[1] 
        } elseif ($tokenData.unique_name -match '@(.+)$') {
            $tenantInfo.TenantDomain = $matches[1] 
        }

        return $tenantInfo
    } catch {
        Write-Log "Error parsing JWT token: $($_.Exception.Message)" -Level "Warning"
        return $null
    }
}

#endregion

#region Main Script

try {
    Write-Log "========================================" -Level "Info"
    Write-Log "Intune App Upload Script" -Level "Info"
    Write-Log "========================================" -Level "Info"

    # Normalize bearer token
    $BearerToken = $BearerToken.Trim()
    if ($BearerToken.StartsWith("Bearer ", [System.StringComparison]::OrdinalIgnoreCase)) {
        $BearerToken = $BearerToken.Substring(7).Trim()
    }

    # Validate token and get tenant info
    $tenantInfo = Get-JwtTenantInfo -JwtToken $BearerToken
    if (-not $tenantInfo -or -not $tenantInfo.IsValid) {
        throw "Invalid Bearer token provided. Please ensure you have a valid Microsoft Graph API token."
    }
    Write-Log "Connected to tenant: $($tenantInfo.TenantDomain)" -Level "Success"
    if ($tenantInfo.UserPrincipalName) {
        Write-Log "User: $($tenantInfo.UserPrincipalName)"
    }

    # Find .intunewin file
    $intuneWinFile = Get-ChildItem -Path $PackagePath -Filter "*.intunewin" | Select-Object -First 1
    if (-not $intuneWinFile) {
        throw "No .intunewin file found in: $PackagePath"
    }
    Write-Log "Found package: $($intuneWinFile.Name)"

    # Prepare headers
    $headers = @{
        'Authorization' = "Bearer $BearerToken"
        'Content-Type'  = 'application/json'
        'Accept'        = 'application/json'
    }

    # Build install/uninstall commands
    $installCommand = "Invoke-AppDeployToolkit.exe -DeploymentType `"Install`" -DeployMode `"$InstallMode`""
    $uninstallCommand = "Invoke-AppDeployToolkit.exe -DeploymentType `"Uninstall`" -DeployMode `"$UninstallMode`""

    Write-Log "Install command: $installCommand"
    Write-Log "Uninstall command: $uninstallCommand"

    # Extract detection.xml from .intunewin
    Write-Log "Extracting metadata from .intunewin file..."
    $detectionXmlData = Get-IntuneWinDetectionXml -IntuneWinPath $intuneWinFile.FullName
    if (-not $detectionXmlData) {
        throw "Failed to extract detection.xml from .intunewin file"
    }

    # Extract encrypted content file
    $encryptedContentFile = Get-IntuneWinContentFile -IntuneWinPath $intuneWinFile.FullName -DetectionXml $detectionXmlData
    if (-not $encryptedContentFile) {
        throw "Failed to extract encrypted content file from .intunewin package"
    }

    # Prepare detection rule - prefer MSI product code when available, otherwise file detection; PowerShell script is explicit opt-in
    $detectionRule = @{}

    $msiProductCode = $null
    if ($detectionXmlData.ApplicationInfo.MsiProductCode) {
        $msiProductCode = $detectionXmlData.ApplicationInfo.MsiProductCode.Trim()
    }

    # Detection parameter processing - supports MSI product codes, file paths, or PowerShell scripts
    if ($Detection) {
        # Check if Detection is an MSI product code GUID (e.g., "{D8C5D09E-1730-3737-80DD-023A666CAF5B}")
        if ($Detection -match '^\{?[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}\}?$') {
            # Ensure GUID has curly braces
            $productCodeGuid = $Detection
            if (-not $productCodeGuid.StartsWith('{')) {
                $productCodeGuid = "{$productCodeGuid}"
            }
            if (-not $productCodeGuid.EndsWith('}')) {
                $productCodeGuid = "$productCodeGuid}"
            }
            Write-Log "Using MSI product code detection: $productCodeGuid"
            $detectionRule = @{
                '@odata.type'          = '#microsoft.graph.win32LobAppProductCodeRule'
                productCode            = $productCodeGuid
                productVersionOperator = 'notConfigured'
                productVersion         = $null
            }
        }
        # Check if Detection is a PowerShell script file
        elseif ((Test-Path $Detection) -and $Detection -like "*.ps1") {
            Write-Log "Using PowerShell detection script: $Detection"
            $scriptContent = Get-Content $Detection -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($scriptContent)) {
                $scriptContentBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($scriptContent))
                $detectionRule = @{
                    '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptRule'
                    enforceSignatureCheck = $false
                    runAs32Bit            = $false
                    scriptContent         = $scriptContentBase64
                }
            } else {
                Write-Log "Detection script is empty, falling back to automatic detection" -Level "Warning"
            }
        }
        # Check if Detection is a file path (e.g., "C:\Program Files\App\file.exe")
        elseif ($Detection -match '^[A-Za-z]:\\|^%[A-Za-z]+%') {
            Write-Log "Using file path detection: $Detection"
            
            # Parse the file path to extract path and filename
            $detectionPath = [System.IO.Path]::GetDirectoryName($Detection)
            $detectionFileName = [System.IO.Path]::GetFileName($Detection)
            
            $detectionRule = @{
                '@odata.type'        = '#microsoft.graph.win32LobAppFileSystemRule'
                ruleType             = 'detection'
                operator             = 'notConfigured'
                check32BitOn64System = $false
                operationType        = 'exists'
                comparisonValue      = $null
                fileOrFolderName     = $detectionFileName
                path                 = $detectionPath
            }
        } else {
            Write-Log "Detection parameter format not recognized, using default detection" -Level "Warning"
        }
    }

    # MSI product code detection (preferred when available and no detection specified)
    if (-not $detectionRule -and $msiProductCode) {
        Write-Log "Using MSI product code detection: $msiProductCode"
        $detectionRule = @{
            '@odata.type'          = '#microsoft.graph.win32LobAppProductCodeRule'
            productCode            = $msiProductCode
            productVersionOperator = 'notConfigured'
            productVersion         = $null
        }
    }

    # Default file system detection if nothing else selected
    if (-not $detectionRule) {
        Write-Log "Using default file detection rule (recommended)"
        $detectionRule = @{
            '@odata.type'        = '#microsoft.graph.win32LobAppFileSystemRule'
            ruleType             = 'detection'
            operator             = 'notConfigured'
            check32BitOn64System = $false
            operationType        = 'exists'
            comparisonValue      = $null
            fileOrFolderName     = $AppName
            path                 = '%ProgramFiles%'
        }
    }

    # Prepare description
    if ([string]::IsNullOrWhiteSpace($Description)) {
        $Description = "Application $AppName. Deployed via Intune using PowerShell Application Deployment Toolkit (PSADT)."
    }

    # Notes field - includes WinGet AppID, saiwParams, and UI Language
    $notesParts = @("Winget:$AppId")
    if (-not [string]::IsNullOrWhiteSpace($InstallParams)) {
        $notesParts += "Params:$InstallParams"
    }
    if (-not [string]::IsNullOrWhiteSpace($UILanguage)) {
        $notesParts += "UILang:$UILanguage"
    }
    $notesParts += "Prod"
    $notes = $notesParts -join ';'

    # Load icon if provided
    $iconData = $null
    if ($IconPath -and (Test-Path $IconPath)) {
        $iconFileName = [System.IO.Path]::GetFileName($IconPath)
        $sanitizedAppName = ($AppName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
        $sanitizedIconName = ($iconFileName -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()

        if ($iconFileName -in @('AppIcon.png', 'banner.classic.png')) {
            Write-Log "Skipping placeholder icon ($iconFileName); not used for upload" -Level "Warning"
        } elseif (-not $sanitizedIconName.Contains($sanitizedAppName)) {
            Write-Log "Icon file '$iconFileName' does not match application name; skipping icon upload" -Level "Warning"
        } else {
            Write-Log "Loading application icon: $IconPath"
            $iconData = [System.IO.File]::ReadAllBytes($IconPath)
        }
    }

    # Build request body
    $requestBody = @{
        "@odata.type"                  = "#microsoft.graph.win32LobApp"
        applicableArchitectures        = "none"
        allowedArchitectures           = $null
        allowAvailableUninstall        = $true
        categories                     = @()
        description                    = $Description
        developer                      = ""
        displayName                    = $AppName
        displayVersion                 = $AppVersion
        fileName                       = $detectionXmlData.ApplicationInfo.FileName
        installCommandLine             = $installCommand
        installExperience              = @{
            "deviceRestartBehavior" = "suppress"
            "maxRunTimeInMinutes"   = 60
            "runAsAccount"          = "system"
        }
        informationUrl                 = ""
        isFeatured                     = $false
        roleScopeTagIds                = @()
        minimumSupportedWindowsRelease = "Windows11_23H2"
        msiInformation                 = $null
        notes                          = $notes
        owner                          = ""
        privacyInformationUrl          = ""
        publisher                      = $AppVendor
        returnCodes                    = $script:returnCodes
        rules                          = @($detectionRule)
        runAs32Bit                     = $false
        setupFilePath                  = $detectionXmlData.ApplicationInfo.SetupFile
        uninstallCommandLine           = $uninstallCommand
        activeInstallScript            = $null
        activeUninstallScript          = $null
    }

    # Add icon if available
    if ($iconData -and $iconData.Length -gt 0) {
        $base64Icon = [Convert]::ToBase64String($iconData)
        $requestBody.largeIcon = @{
            type  = 'image/png'
            value = $base64Icon
        }
        Write-Log "Application icon included"
    }

    # Step 1: Create Win32 app
    Write-Log "Creating Win32 app in Intune..."
    $requestBodyJson = ConvertTo-Json -InputObject $requestBody -Depth 10

    $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps" -Method POST -Headers $headers -Body $requestBodyJson
    $appId = $response.id
    Write-Log "App created with ID: $appId" -Level "Success"

    # Step 2: Create content version
    Write-Log "Creating content version..."
    $contentVersionUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions"
    $contentVersionResponse = Invoke-RestMethod -Uri $contentVersionUri -Method POST -Headers $headers -Body "{}"
    $contentVersionId = $contentVersionResponse.id
    Write-Log "Content version ID: $contentVersionId"

    # Step 3: Create file entry
    Write-Log "Creating file entry..."
    $encryptedFileInfo = Get-Item $encryptedContentFile
    $filePayload = @{
        '@odata.type' = '#microsoft.graph.mobileAppContentFile'
        name          = [System.IO.Path]::GetFileName($detectionXmlData.ApplicationInfo.FileName)
        size          = [int64]$detectionXmlData.ApplicationInfo.UnencryptedContentSize
        sizeEncrypted = $encryptedFileInfo.Length
        manifest      = $null
        isDependency  = $false
    } | ConvertTo-Json -Depth 10

    $fileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions/$contentVersionId/files"
    $fileResponse = Invoke-RestMethod -Uri $fileUri -Method POST -Headers $headers -Body $filePayload
    $fileId = $fileResponse.id
    Write-Log "File entry ID: $fileId"

    # Step 4: Wait for Azure Storage URI
    Write-Log "Waiting for Azure Storage URI..."
    $fileStatusUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions/$contentVersionId/files/$fileId"
    do {
        Start-Sleep -Seconds 3
        $fileStatus = Invoke-RestMethod -Uri $fileStatusUri -Method GET -Headers $headers
    } while ($fileStatus.uploadState -ne 'azureStorageUriRequestSuccess')

    # Step 5: Upload to Azure Storage
    Write-Log "Uploading to Azure Storage..."
    $uploadSuccess = Invoke-AzureStorageUpload -AzureStorageUri $fileStatus.azureStorageUri -FilePath $encryptedContentFile -FileContainerUri $fileStatusUri
    if (-not $uploadSuccess) {
        throw "Azure Storage upload failed"
    }

    # Step 6: Commit file
    Write-Log "Committing file with encryption information..."
    $encryptionInfo = @{
        encryptionKey        = $detectionXmlData.ApplicationInfo.EncryptionInfo.EncryptionKey
        macKey               = $detectionXmlData.ApplicationInfo.EncryptionInfo.macKey
        initializationVector = $detectionXmlData.ApplicationInfo.EncryptionInfo.initializationVector
        mac                  = $detectionXmlData.ApplicationInfo.EncryptionInfo.mac
        profileIdentifier    = "ProfileVersion1"
        fileDigest           = $detectionXmlData.ApplicationInfo.EncryptionInfo.fileDigest
        fileDigestAlgorithm  = $detectionXmlData.ApplicationInfo.EncryptionInfo.fileDigestAlgorithm
    }

    $commitPayload = @{ fileEncryptionInfo = $encryptionInfo } | ConvertTo-Json -Depth 10
    $commitUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$appId/microsoft.graph.win32LobApp/contentVersions/$contentVersionId/files/$fileId/commit"
    Invoke-RestMethod -Uri $commitUri -Method POST -Headers $headers -Body $commitPayload

    # Step 7: Wait for commit
    Write-Log "Waiting for file commit to complete..."
    do {
        Start-Sleep -Seconds 5
        $fileCommitStatus = Invoke-RestMethod -Uri $fileStatusUri -Method GET -Headers $headers
    } while ($fileCommitStatus.uploadState -eq 'commitFilePending')

    if ($fileCommitStatus.uploadState -eq 'commitFileFailed') {
        throw "File commit failed"
    }

    # Step 8: Commit the app
    Write-Log "Committing application..."
    $commitAppPayload = @{
        '@odata.type'           = '#microsoft.graph.win32LobApp'
        committedContentVersion = $contentVersionId
    } | ConvertTo-Json -Depth 2

    $commitAppUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$appId"
    Invoke-RestMethod -Uri $commitAppUri -Method PATCH -Headers $headers -Body $commitAppPayload

    # Cleanup
    try {
        Remove-Item $encryptedContentFile -Force -ErrorAction SilentlyContinue
        if ($script:tempExtractionDir -and (Test-Path $script:tempExtractionDir)) {
            Remove-Item $script:tempExtractionDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { 
    }

    Write-Log "========================================" -Level "Success"
    Write-Log "Upload completed successfully!" -Level "Success"
    Write-Log "Intune App ID: $appId" -Level "Success"
    Write-Log "Tenant: $($tenantInfo.TenantDomain)" -Level "Success"
    Write-Log "========================================" -Level "Success"

    # Output result for agent to parse
    return @{
        Success = $true
        AppId   = $appId
        AppName = $AppName
        Tenant  = $tenantInfo.TenantDomain
        Message = "App uploaded successfully"
    }
} catch {
    $errorMessage = $_.Exception.Message
    
    # Try to get detailed error response
    try {
        if ($_.Exception.Response) {
            $responseStream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($responseStream)
            $responseBody = $reader.ReadToEnd()
            $reader.Close()
            
            try {
                $errorJson = ConvertFrom-Json $responseBody
                if ($errorJson.error.message) {
                    $errorMessage = "API Error: $($errorJson.error.message)"
                }
            } catch { 
            }
        }
    } catch { 
    }

    # Cleanup on error
    try {
        if ($encryptedContentFile -and (Test-Path $encryptedContentFile)) {
            Remove-Item $encryptedContentFile -Force -ErrorAction SilentlyContinue
        }
        if ($script:tempExtractionDir -and (Test-Path $script:tempExtractionDir)) {
            Remove-Item $script:tempExtractionDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { 
    }

    Write-Log "========================================" -Level "Error"
    Write-Log "Upload failed: $errorMessage" -Level "Error"
    Write-Log "========================================" -Level "Error"

    return @{
        Success = $false
        AppId   = $null
        Message = $errorMessage
    }
}

#endregion
