# Using CLI for Microsoft 365 and PowerShell to report on SharePoint Document Library files containing PII

Author: [Joseph Velliah](https://sprider.blog/report-on-sharepoint-files-containing-pii)

Personally Identifiable Information (PII) is information that, when used alone or with other relevant data, can identify an individual. Sensitive personally identifiable information can include full name, SSN, driver’s license, financial information and medical records. As PII can be used to identify an individual, signify a major threat to companies. If breached, this information can lead to lawsuits and can damage company’s trustworthiness.

This script can be used to scan and report on SharePoint Document Library files (specified file types) containing common PII. This script searches and reports only on social security numbers and credit card numbers. Feel free to add additional PII patterns.

Thank you, [Sam Boutros](https://www.linkedin.com/in/sam-boutros-powershell/), for your valuable article on this topic.

Prerequisites:

- [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
- SharePoint Online Document Library files containing dummy/real common PII data in txt, csv, aspx file formats
- [Here](https://dlptest.com/sample-data/) is some test data and test files that can be used to test.

```powershell tab="PowerShell Core"
# This script is created to support only the following file extensions.
$supportedFileExtensions = "txt", "csv", "aspx"
$spolHostName = "https://tenant-name.sharepoint.com"
$spolSiteRelativeUrl = "/sites/site-name"
$spolDocLibTitle = "library-title"
$resultDir = "Output"

$Patterns = @()
$Pattern0 = 'Credit Card - Visa', '(4\d{3}[-| ]\d{4}[-| ]\d{4}[-| ]\d{4})|(4\d{15})', 'Start with a 4 and have 16 digits, may be split as xxxx-xxxx-xxxx-xxxx by dashes or spaces'
$Pattern1 = 'Credit Card - MasterCard', '(5[1-5]\d{14})|(5[1-5]\d{2}[-| ]\d{4}[-| ]\d{4}[-| ]\d{4})', 'Starts with 51-55 and have 16 digits, may be split as xxxx-xxxx-xxxx-xxxx by dashes or spaces'
$Pattern2 = 'Credit Card - Amex', '(3[47]\d{13})|(3[47]\d{2}[-| ]\d{6}[-| ]\d{5})', 'Starts with 34 or 37 and have 15 digits, may be split as xxxx-xxxxxx-xxxxx by dashes or spaces'
$Pattern3 = 'Credit Card - DinersClub', '(3(?:0[0-5]|[68]\d)\d{11})|(3(?:0[0-5]|[68]\d)\d[-| ]\d{6}[-| ]\d{4})', 'Starts with 300-305, or 36-38 and have 14 digits, may be split as xxxx-xxxxxx-xxxx by dashes or spaces'
$Pattern4 = 'Credit Card - Discover', '(6(?:011|5\d{2})\d{12})|(6(?:011|5\d{2})[-| ]\d{4}[-| ]\d{4}[-| ]\d{4})', 'Start with 6011 or 65 and have 16 digits, may be split as xxxx-xxxx-xxxx-xxxx by dashes or spaces'
$Pattern5 = 'Credit Card - JCB', '((?:2131|1800|35\d{3})\d{11})|((?:2131|1800|35\d{2})[-| ]\d{4}[-| ]\d{4}[-| ]\d{3}[\d| ])', 'Start with 2131 or 1800 and have 15 digits) or (Start with 35 and have 16 digits'
$Pattern6 = 'Social Security Number', '(\d{3}[-| ]\d{2}[-| ]\d{4})|(\d{9})', '9 digits, may be split as xxx-xx-xxxx by dashes or spaces'
$Patterns = $Pattern0, $Pattern1, $Pattern2, $Pattern3, $Pattern4, $Pattern5, $Pattern6

$executionDir = $PSScriptRoot
$outputDir = "$executionDir/$resultDir"
$outputFilePath = "$outputDir/$(get-date -f yyyyMMdd-HHmmss)-PIIFindings.csv"

if (-not (Test-Path -Path "$outputDir" -PathType Container)) {
    Write-Host "Creating $outputDir folder..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path "$outputDir"
}

$spolSiteUrl = $spolHostName + $spolSiteRelativeUrl
$spolLibItems = m365 spo listitem list --webUrl $spolSiteUrl --title $spolDocLibTitle --fields 'FileRef,FileLeafRef,File_x0020_Type' --filter "FSObjType eq 0" -o json | ConvertFrom-Json

if ($spolLibItems.Count -gt 0) {
    $spFilePIIFindings = @()

    ForEach ($spolLibItem in $spolLibItems) {
        $spolLibFileRelativeUrl = $spolLibItem.FileRef
        $spolLibFileExtension = $spolLibItem.File_x0020_Type

        if ($supportedFileExtensions.Contains($spolLibFileExtension)) {
            Write-Host "Processing file: $spolLibFileRelativeUrl" -ForegroundColor Magenta

            try {
                $fileContent = m365 spo file get --webUrl $spolSiteUrl --url $spolLibFileRelativeUrl --asString

                for ($j = 0; $j -lt $fileContent.Count; $j++) {
                    For ($i = 0; $i -lt $Patterns.Count; $i++) {
                        if ($fileContent[$j] -match $Patterns[$i][1]) {
                            $spFilePIIFinding = New-Object -TypeName PSObject
                            $spFilePIIFinding | Add-Member -MemberType NoteProperty -Name "FileRelativeUrl" -Value $spolLibFileRelativeUrl
                            $spFilePIIFinding | Add-Member -MemberType NoteProperty -Name "Pattern" -Value $Patterns[$i][0]
                            $spFilePIIFinding | Add-Member -MemberType NoteProperty -Name "Line" -Value $j
                            $spFilePIIFinding | Add-Member -MemberType NoteProperty -Name "Content" -Value $fileContent[$j]
                            $spFilePIIFindings += $spFilePIIFinding
                        }
                    }
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

    If ($spFilePIIFindings.Length -gt 0) {
        $spFilePIIFindings | Export-Csv -Path "$outputFilePath" -NoTypeInformation
        Write-Host "Open $outputFilePath to review PII Findings report." -ForegroundColor Green
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
- PII
