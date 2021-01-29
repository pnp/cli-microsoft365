# Export conversations from Microsoft Teams Channels

Author: [Joseph Velliah](https://sprider.blog/export-conversations-from-microsoft-teams)

## Problem Statement

We utilize Teams during incidents and create channels for each. We would like to be able to export conversation history.

- Right now the only option we have is to go through Security & Compliance.
- Teams usage is growing in every organization and it would soon become unreasonably to only have Administrators be the ones doing exports of channels for all the Teams.

### Solution

This script uses Microsoft 365 CLI to export the conversations from Microsoft Teams Channels.

!!! notes

- Commands m365 teams message list and reply list are based on an API that is currently in preview and is subject to change once the API reached general availability.
- You can only retrieve a message from a Microsoft Teams team if you are a member of that team.

```powershell tab="PowerShell Core"
function  Get-Teams {
    $teams = m365 teams team list -o json | ConvertFrom-Json -AsHashtable
    return $teams
}
function  Get-Channels {
    param (
        [Parameter(Mandatory = $true)] [string] $teamId
    )
    $channels = m365 teams channel list --teamId $teamId -o json | ConvertFrom-Json -AsHashtable
    return $channels
}
function  Get-Messages {
    param (
        [Parameter(Mandatory = $true)] [string] $teamId,
        [Parameter(Mandatory = $true)] [string] $channelId
    )
    $messages = m365 teams message list --teamId $teamId --channelId $channelId -o json | ConvertFrom-Json -AsHashtable
    return $messages
}
function  Get-MessageReplies {
    param (
        [Parameter(Mandatory = $true)] [string] $teamId,
        [Parameter(Mandatory = $true)] [string] $channelId,
        [Parameter(Mandatory = $true)] [string] $messageId
    )

    $messageReplies = m365 teams message reply list --teamId $teamId --channelId $channelId --messageId $messageId -o json | ConvertFrom-Json -AsHashtable
    return $messageReplies
}

$teamsCollection = [System.Collections.ArrayList]@()
$teams = Get-Teams
$teamsCount = $teams.Length
Write-Host "$teamsCount Team/s found" -ForegroundColor Magenta
foreach ($team in $teams) {
    $channelsCollection = [System.Collections.ArrayList]@()
    $channels = Get-Channels $team.id
    $teamDisplayName = $team.displayName
    $channelsCount = $channels.Length
    Write-Host "    $channelsCount Channel/s found in Team $teamDisplayName" -ForegroundColor Blue
    foreach ($channel in $channels) {
        $channelDisplayName = $channel.displayName
        Write-Host "        Collecting conversation details from Channel $channelDisplayName" -ForegroundColor Gray
        $messages = Get-Messages $team.id $channel.id
        $messagesCollection = [System.Collections.ArrayList]@()
        foreach ($message in $messages) {
            $messageReplies = Get-MessageReplies $team.id $channel.id $message.id
            $messageDetails = $message
            [void]$messageDetails.Add("replies", $messageReplies)
            [void]$messagesCollection.Add($messageDetails)
        }
        $channelDetails = $channel
        [void]$channelDetails.Add("messages", $messagesCollection)
        [void]$channelsCollection.Add($channelDetails)
        Write-Host "        Completed" -ForegroundColor DarkGray
    }
    $teamDetails = $team
    [void]$teamDetails.Add("channels", $channelsCollection)
    [void]$teamsCollection.Add($teamDetails)
}
$output = @{}
[void]$output.Add("teams", $teamsCollection)

Write-Host "Creating output file" -ForegroundColor Green
$executionDir = $PSScriptRoot
$outputFilePath = "$executionDir/$(get-date -f yyyyMMdd-HHmmss).json"
$output | ConvertTo-Json -Depth 10 | Out-File $outputFilePath
Write-Host "Open $outputFilePath file to review your output" -ForegroundColor DarkGreen
```

Keywords:

- Microsoft Teams
- PowerShell
