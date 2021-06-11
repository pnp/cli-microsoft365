# Flow runs day summary report

Author: [Adam](https://github.com/Adam-it)

This script creates a report of all flow runs from current day and sends the report as an adaptive card to the provided url

=== "PowerShell"

    ```powershell
    $m365Status = m365 status
    if ($m365Status -eq "Logged Out") {
        m365 login
    }

    $environment = 'Default-2942bb31-1d49-4da6-8d3d-d0f9e1141234'
    $adaptiveCard = '{\"type\":\"AdaptiveCard\",\"body\":[{\"type\":\"TextBlock\",\"size\":\"Medium\",\"weight\":\"Bolder\",\"text\":\"${title}\"},{\"type\":\"TextBlock\",\"text\":\"${description}\",\"wrap\":true}],\"$schema\":\"http://adaptivecards.io/schemas/adaptive-card.json\",\"version\":\"1.3\"}'
    $webhook = 'https://contoso.webhook.office.com/webhookb2/1204eba2-061c-4442-9696-2a725cb2d094@2942bb31-1d49-4da6-8d3d-d0f9e1141486/IncomingWebhook/6e54c3958bde444e96fec9ecad356993/be11f523-2a4d-4eae-9d42-277410893c41'

    $flows = m365 flow list --environment $environment --output json
    $flows = $flows | ConvertFrom-Json
    $currentDayDate = Get-Date
    $previousDayDate = (Get-Date).AddDays(-1)

    $adaptiveCardDescription = ""
    foreach ($flow in $flows) 
    {
        $flowRuns = m365 flow run list --environment $environment --flow $flow.name --output json
        $flowRuns = $flowRuns | ConvertFrom-Json

        $displayName = $flow.displayName
        $id = $flow.name

        $todayRuns = $flowRuns.Where({[DateTime]$_.properties.endTime -le $currentDayDate -and [DateTime]$_.properties.endTime -gt $previousDayDate})
        
        $todayRunsCount = 0
        $todaySuccessRunsCount = 0
        $todayFailedRunsCount = 0
        if($todayRuns.Count -gt 0)
        {
            $todaySuccessRuns = $todayRuns.Where({$_.status -eq 'Succeeded'})
            $todaySuccessRunsCount = $todaySuccessRuns.Count

            $todayFailedRuns = $todayRuns.Where({$_.status -eq 'Failed'})
            $todayFailedRunsCount = $todayFailedRuns.Count

            $todayRunsCount = $todayRuns.Count
        }

        Write-Host "$displayName -> Runs: $todayRunsCount , Succeeded: $todaySuccessRunsCount , Failed: $todayFailedRunsCount"
        $adaptiveCardDescription = $adaptiveCardDescription + "\r- [$displayName](https://us.flow.microsoft.com/manage/environments/$environment/flows/$id/details) -> Runs: $todayRunsCount , Succeeded: $todaySuccessRunsCount , Failed: $todayFailedRunsCount"
    }

    $today = Get-Date -Format "MM/dd/yyyy"
    $cardData = '{\"title\": \"Flows summary - ' + $today + '\" ,\"description\":\"' + $adaptiveCardDescription + '\"}'

    m365 adaptivecard send --url $webhook --card $adaptiveCard --cardData $cardData
    ```

Keywords:

- Flow
- Power Automate
- Adaptive Card
