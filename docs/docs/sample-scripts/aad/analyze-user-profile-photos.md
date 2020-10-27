# Analyze User Profile Photos using Azure Computer Vision API

Author: [Joseph Velliah](https://sprider.blog/analyze-microsoft-365-user-profile-photos-using-azure-computer-vision-api)

With 2020 and the onset of a new decade, there has been a dramatic shift in the professional working paradigm. With more people working from home, video meetings, virtual discussion rooms and other AI tools becoming the norm, technology is connecting us across oceans and continents. Yet, with ‘great power comes great responsibility’ and with ‘the great power of digital connection comes increasing impersonality.’

As the adage goes ‘First impressions are the best impressions’, now is the perfect time to make a lasting and valuable first impression with your profile photos.

This script uses Azure Cognitive Service API and Microsoft 365 CLI to analyze user profile pictures and assess whether they meet the standards placed by the organization.

This script can be customized to ban content within an org channel or collaboration network where employees post pictures, memes, etc.

Prerequisites

- [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
- Microsoft 365 users
- [Computer Vision API](https://azure.microsoft.com/en-us/services/cognitive-services/computer-vision/) instance and API key

Note: If you don't already have an [Azure Cognitive Services instance and key](https://azure.microsoft.com/en-us/try/cognitive-services/), create a cognitive service instance and get API key from there.

I would like to thank [Hugo Bernier](https://github.com/pnp/sp-dev-fx-webparts/tree/master/samples/react-smart-profile-photo-editor), for the sample SharePoint Framework WebPart.

There is no better time than the present to get your profile photo game on point. So, get working on it.

```powershell tab="PowerShell Core"
$resultDir = "Output"
$azureVisionApiInstance = "azure-vision-api-instance-name"
$azureVisionApiKey = "azure-vision-api-key"

$photoRequirements = @{
    requirePortrait   = $false 
    allowClipart      = $true
    allowLinedrawing  = $true
    allowAdult        = $false
    allowRacy         = $false
    allowGory         = $false
    photoRequirements = @{
        requirePortrait   = $false
        allowClipart      = $true
        allowLinedrawing  = $true
        allowAdult        = $false
        allowRacy         = $false
        allowGory         = $false
        forbiddenKeywords = @("cartoon", `
                "animal", `
                "nude", `
                "child", `
                "people", `
                "group", `
                "family", `
                "several", `
                "crowd", `
                "food", `
                "restaurant", `
                "train", `
                "bus", `
                "car", `
                "airplane", `
                "vehicle", `
                "platform", `
                "station", `
                "standing", `
                "flying", `
                "suitcase", `
                "screenshot", `
                "newspaper", `
                "typography", `
                "font", `
                "document", `
                "sport")
    }
}

$requiredProfileProperties = "id,displayName,mail"
$global:analysisOutcomes = @()

$executionDir = $PSScriptRoot
$outputDir = "$executionDir/$resultDir"
$outputFilePath = "$outputDir/$(get-date -f yyyyMMdd-HHmmss)-scan-profile-pictures-outcome.csv"

if (-not (Test-Path -Path "$outputDir" -PathType Container)) {
    Write-Host "Creating $outputDir folder..."
    New-Item -ItemType Directory -Path "$outputDir"
}

function AddAnalysisOutcome {
    param (
        [Parameter(Mandatory = $false)] [string] $UserId,
        [Parameter(Mandatory = $false)] [string] $UserMail,
        [Parameter(Mandatory = $false)] [bool] $IsPortraitValid,
        [Parameter(Mandatory = $false)] [bool] $IsOnlyOnePersonValid,
        [Parameter(Mandatory = $false)] [bool] $IsClipartValid,
        [Parameter(Mandatory = $false)] [bool] $IsLineDrawingValid,
        [Parameter(Mandatory = $false)] [bool] $IsAdultValid,
        [Parameter(Mandatory = $false)] [bool] $IsRacyValid,
        [Parameter(Mandatory = $false)] [bool] $IsGoryValid,
        [Parameter(Mandatory = $false)] [bool] $IsCelebrity,
        [Parameter(Mandatory = $false)] [bool] $IsForbiddenKeywordExist,
        [Parameter(Mandatory = $false)] [bool] $IsValidProfilePhoto,
        [Parameter(Mandatory = $false)] [string] $Notes
    )

    $analysisOutcome = New-Object -TypeName PSObject

    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "UserId" -Value $UserId
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "UserMail" -Value $UserMail
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsPortraitValid" -Value $IsPortraitValid
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsOnlyOnePersonValid" -Value $IsOnlyOnePersonValid
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsClipartValid" -Value $IsClipartValid
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsLineDrawingValid" -Value $IsLineDrawingValid
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsAdultValid" -Value $IsAdultValid
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsRacyValid" -Value $IsRacyValid
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsGoryValid" -Value $IsGoryValid
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsCelebrity" -Value $IsCelebrity
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsForbiddenKeywordExist" -Value $IsForbiddenKeywordExist
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "IsValidProfilePhoto" -Value $IsValidProfilePhoto
    $analysisOutcome | Add-Member -MemberType NoteProperty -Name "Notes" -Value $Notes

    $global:analysisOutcomes += $analysisOutcome
}

$users = m365 aad user list --properties $requiredProfileProperties -o json | ConvertFrom-Json -AsHashtable

foreach ($user in $users) {

    try {
        $userId = $user.id
        $userMail = $user.mail

        try {
            $token = m365 util accesstoken get --resource https: / / graph.microsoft.com --new

            try {
                $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
                $headers.Add("Content-Type", "image/jpg")
                $headers.Add("Authorization", "Bearer $token")
                $userPhoto = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$userId/photo/`$value" -Headers $headers)

                If ($userPhoto) {

                    try {

                        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
                        $headers.Add("Content-Type", "application/json")
                        $headers.Add("Ocp-Apim-Subscription-Key", $azureVisionApiKey)

                        $analysis = (Invoke-RestMethod -Uri ("https://$azureVisionApiInstance.cognitiveservices.azure.com/vision/v3.1/analyze?visualFeatures=Categories,Adult,Tags,Description,Faces,Color,ImageType,Objects&details=Celebrities&language=en") `
                                -Headers $headers `
                                -Body ($userPhoto) `
                                -ContentType "application/octet-stream" `
                                -Method "Post");

                        if ($analysis) {
                            $analysisData = $analysis | ConvertFrom-Json -AsHashtable
                            $isPortrait = $analysisData.categories.Length -gt 0 ? ($analysisData.categories | Where-Object { $_.name -eq 'people_portrait' }).Length -gt 0  ? $true : $false : $false
                            $isPortraitValid = $photoRequirements.requirePortrait ? $isPortrait : $true
                            $isOnlyOnePersonValid = $analysisData.faces.Length -eq 1 ? $true : $false
                            $isClipartValid = $analysisData.imageType.clipArtType -eq 0 ? $true : $false
                            $isLineDrawingValid = $analysisData.imageType.lineDrawingType -eq 0 ? $true : $false
                            $isAdultValid = $photoRequirements.allowAdult ? $true : !$analysisData.adult.isAdultContent
                            $isRacyValid = $photoRequirements.allowRacy ? $true : !$analysisData.adult.isRacyContent
                            $isGoryValid = $photoRequirements.allowGory ? $true : !$analysisData.adult.isGoryContent
                            $isCelebrity = ($analysisData.categories | Where-Object { $_.detail.celebrities.Length -gt 0 }).Length -gt 0 ? $true : $false

                            $invalidKeywords = @()

                            foreach ($forbiddenKeyword in $photoRequirements.forbiddenKeywords) {
                                $isForbiddenKeywordExist = ($analysisData.tags | Where-Object { $_.name -eq $forbiddenKeyword }).Length -gt 0 ? $true : $false

                                if ($isForbiddenKeywordExist) {
                                    $invalidKeyword = New-Object -TypeName PSObject
                                    $invalidKeyword | Add-Member -MemberType NoteProperty -Name forbiddenKeyword -Value $forbiddenKeyword
                                    $invalidKeywords += $invalidKeyword
                                }
                            }

                            $isForbiddenKeywordExist = $invalidKeywords.Length -gt 0 ? $true : $false

                            $isValidProfilePhoto = $isPortraitValid `
                                -and $isOnlyOnePersonValid `
                                -and $isClipartValid  `
                                -and $isLineDrawingValid `
                                -and $isAdultValid `
                                -and $isRacyValid `
                                -and $isGoryValid `
                                -and !$isCelebrity `
                                -and !$isForbiddenKeywordExist;

                            AddAnalysisOutcome $userId `
                                $userMail `
                                $isPortraitValid `
                                $isOnlyOnePersonValid `
                                $isClipartValid `
                                $isLineDrawingValid `
                                $isAdultValid `
                                $isRacyValid `
                                $isGoryValid `
                                $isCelebrity `
                                $isForbiddenKeywordExist `
                                $isValidProfilePhoto `
                                "Profile photo available"
                        }
                    }
                    catch {
                        AddAnalysisOutcome $userId `
                            $userMail `
                            $false `
                            $false `
                            $false `
                            $false `
                            $false `
                            $false `
                            $false `
                            $false `
                            $false `
                            $false `
                            "Unable to analyze profile photo"
                    }
                }
            }
            catch {
                AddAnalysisOutcome $userId `
                    $userMail `
                    $false `
                    $false `
                    $false `
                    $false `
                    $false `
                    $false `
                    $false `
                    $false `
                    $false `
                    $false `
                    "Unable to get profile photo"
            }
        }
        catch {
            Write-Host "Unable to get new access token" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Unable to get profile details for this user" -ForegroundColor Red
    }
}

$global:analysisOutcomes | Export-Csv -Path "$outputFilePath" -NoTypeInformation
Write-Host "Open $outputFilePath to review analysis outcomes report."
```

Keywords:

- Azure
- Computer Vision API
- Microsoft 365
- PowerShell
