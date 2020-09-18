# Using CLI for Microsoft 365 and PowerShell to report on SharePoint Document Library files containing Client Id or Secret

Author: [Joseph Velliah](https://sprider.blog/report-on-sharepoint-files-containing-pii)

The big challenge with using APIs that require authentication in your JavaScript is that you’re forced to expose your API credentials to use them. Anyone who knows how to view source or view requests in their browser’s Developer Tools can view those credentials, steal them, and use them to access the API as you.

This script can be used to scan and report on SharePoint Document Library files (specified file types) containing Client Id or Secret. Feel free to add additional patterns to match your requirements.

Prerequisites:

- [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
- SharePoint Online Document Library files(aspx, js, html) containing dummy/real Client Id or Secrets

```powershell tab="PowerShell Core"
# This script is created to support only the following file extensions.
$supportedFileExtensions = "aspx", "js", "html"
$spolHostName = "https://tenant-name.sharepoint.com"
$spolSiteRelativeUrl = "/sites/site-name"
$spolDocLibTitle = "library-title"
$resultDir = "Output"


$Patterns = @()

$Pattern0 = @{
    Title       = "Client ID"
    Regex       = "(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}"
    Description = "36 digits GUID"
}
$Pattern0Obj = New-Object -Type psobject -Property $Pattern0
$Patterns += $Pattern0Obj

$Pattern1 = @{
    Title       = "Client Secret"
    Regex       = "(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*([^a-zA-Z\d\s])).{32,44}"
    Description = "32 to 44 digits Client Secret"
}
$Pattern1Obj = New-Object -Type psobject -Property $Pattern1
$Patterns += $Pattern1Obj

$executionDir = $PSScriptRoot
$outputDir = "$executionDir/$resultDir"
$outputFilePath = "$outputDir/$(get-date -f yyyyMMdd-HHmmss)-Findings.csv"

if (-not (Test-Path -Path "$outputDir" -PathType Container)) {
    Write-Host "Creating $outputDir folder..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path "$outputDir" | Out-Null
}

$spolSiteUrl = $spolHostName + $spolSiteRelativeUrl
$spolLibItems = o365 spo listitem list --webUrl $spolSiteUrl --title $spolDocLibTitle --fields 'FileRef,FileLeafRef,File_x0020_Type' --filter "FSObjType eq 0" -o json | ConvertFrom-Json -AsHashtable

if ($spolLibItems.Count -gt 0) {
    $spFileFindings = @()

    ForEach ($spolLibItem in $spolLibItems) {
        $spolLibFileRelativeUrl = $spolLibItem.FileRef
        $spolLibFileExtension = $spolLibItem.File_x0020_Type

        if ($supportedFileExtensions.Contains($spolLibFileExtension)) {
            Write-Host "Processing file: $spolLibFileRelativeUrl" -ForegroundColor Magenta

            try {
                $fileContent = o365 spo file get --webUrl $spolSiteUrl --url $spolLibFileRelativeUrl --asString

                try {
                    $lines = $fileContent -split [System.Environment]::NewLine
                    for ($j = 0; $j -lt $lines.Count; $j++) {
                        for ($i = 0; $i -lt $Patterns.Count; $i++) {
                            if ($lines[$j] -match $Patterns[$i].Regex) {
                                $spFileFinding = New-Object -TypeName PSObject
                                $spFileFinding | Add-Member -MemberType NoteProperty -Name "FileRelativeUrl" -Value $spolLibFileRelativeUrl
                                $spFileFinding | Add-Member -MemberType NoteProperty -Name "Pattern" -Value $Patterns[$i].Title
                                $spFileFinding | Add-Member -MemberType NoteProperty -Name "Line" -Value $j
                                $spFileFinding | Add-Member -MemberType NoteProperty -Name "Content" -Value $lines[$j]
                                $spFileFindings += $spFileFinding
                            }
                        }
                    }
                }
                catch {
                    Write-Host "Unable to validate the patterns: $_" -ForegroundColor Red
                }
            }
            catch {
                Write-Host "Unable to read file: $spolLibFileRelativeUrl" -ForegroundColor Red
            }
        }
        else {
            Write-Host "File type $spolLibFileExtension is not supported to scan" -ForegroundColor Yellow
        }
    }

    if ($spFileFindings.Length -gt 0) {
        $spFileFindings | Export-Csv -Path "$outputFilePath" -NoTypeInformation
        Write-Host "Open $outputFilePath to review Findings report." -ForegroundColor Green
    }
    else {
        Write-Host "There are no findings" -ForegroundColor Yellow
    }
}
else {
    Write-Host "No files in $spolDocLibTitle library" -ForegroundColor Yellow
}
```

Keywords:

- CLI for Microsoft 365
- PowerShell
- SharePoint Online
- Governance
