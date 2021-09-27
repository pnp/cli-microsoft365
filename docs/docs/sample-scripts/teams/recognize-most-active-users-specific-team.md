# Recognize most active users for a specific Team

Author: [Albert-Jan Schot](https://www.cloudappie.nl/recognize-active-team-members-cli-microsoft-365/)

Retrieves all activities for a specific Microsoft Teams Team and shares the top 3 contributors based on their score as an adaptive card to the specified webhook url.

```powershell tab="PowerShell"
$teamId = "<PUTYOURTEAMIDHERE>"
$webhookUrl = "<PUTYOURURLHERE>"
# You can get a delta of messages since the last 'n' days. Currently set to seven. You can go back a maximum of 8 months.
$date = (get-date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")

$channels = m365 teams channel list --teamId $teamId --output json | ConvertFrom-Json
$results = @()
$scoreResults = @()

$channelCounter = 0;

foreach ($channel in $channels) {

    $channelCounter++;
    Write-Output "Processing channel... $channelCounter/$($channels.Length)"

    $messages = m365 teams message list --teamId $teamId --channelId $channel.id --since $date --output json | ConvertFrom-Json

    $messageCounter = 0;

    foreach ($message in $messages) {
        $messageCounter++
        Write-Output "Processing message ... $messageCounter/$($messages.Length)"

        # Skip messages that are created with an application (bots / adaptive cards)
        if ($null -ne $message.from.user.id) {
            $results += [pscustomobject][ordered]@{
                Type       = "Post"
                Details    = $message.reactionType
                UserId     = $message.from.user.id
                HasSubject = $($null -ne $message.subject)
            }
        }

        # Process all likes and comments on the initial message
        foreach ($reaction in $message.reactions) {
            $results += [pscustomobject][ordered]@{
                Type    = "Reaction"
                Details = $reaction.reactionType
                UserId  = $reaction.user.user.id
            }
        }

        $replies = m365 teams message reply list --teamId $teamId --channelId $channel.id --messageId $message.Id --output json | ConvertFrom-Json

        foreach ($reply in $replies) {
            # Skip replies that are created with an application (bots)
            if ($null -ne $message.from.user.id) {
                $results += [pscustomobject][ordered]@{
                    Type   = "Reply"
                    UserId = $reply.from.user.id
                }
            }

            # Process all likes and comments on the reply message
            foreach ($reaction in $reply.reactions) {
                $results += [pscustomobject][ordered]@{
                    Type    = "Reaction"
                    Details = $reaction.reactionType
                    UserId  = $reaction.user.user.id
                }
            }

        }
    }
}

# Group the results per user
$resultsGrouped = $results | Group-Object -Property UserId

#Score per user
foreach ($teamsUser in $resultsGrouped) {
    $user = m365 aad user get --id $teamsUser.Name --output json | ConvertFrom-Json

    # Count points
    # Each  post is two points, 1 extra point awarded for each Post with Subject
    # Each reply 1 and each reaction 0.5
    $score = (($teamsUser.Group | Where-Object { $_.Type -eq "Post" }).Count * 2)
    $score += (($teamsUser.Group | Where-Object { $_.HasSubject }).Count)
    $score += ($teamsUser.Group | Where-Object { $_.Type -eq "Reply" }).Count
    $score += (($teamsUser.Group | Where-Object { $_.Type -eq "Reaction" }).Count / 2)

    $scoreResults += [pscustomobject][ordered]@{
        DisplayName       = $user.displayName
        UserPrincipalName = $user.userPrincipalName
        Score             = $score;
    }
}

# Sort our score report based on the score
$scoreResults = $scoreResults | Sort-Object { $_.score } -Descending

#Construct adaptive card
$title = "🏆 Most active team members"
$scoreJson = '{   \"title\": \"🥇 '+$($scoreResults[0].DisplayName)+'\",   \"value\": \"' + $($scoreResults[0].score) + '\"   }'

if($scoreResults[1]){
    $scoreJson += ',{   \"title\": \"🥈 '+$($scoreResults[1].DisplayName)+'\",   \"value\": \"' + $($scoreResults[1].score) + '\"   }'
}
if($scoreResults[2]){
    $scoreJson += ',{   \"title\": \"🥉 '+$($scoreResults[2].DisplayName)+'\",   \"value\": \"' + $($scoreResults[2].score) + '\"   }'
}

$card = '{ \"type\": \"AdaptiveCard\", \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\", \"version\": \"1.2\", \"body\": [  {  \"type\": \"TextBlock\",  \"text\": \"' + $($title) + '\",  \"wrap\": true,  \"size\": \"Medium\",  \"weight\": \"Bolder\",  \"color\": \"Attention\"  },  {  \"type\": \"TextBlock\",  \"wrap\": true,  \"text\": \"Week ' + $(get-date -UFormat %V) + '\",  \"fontType\": \"Default\",  \"size\": \"Small\",  \"weight\": \"Lighter\",  \"isSubtle\": true  },  {  \"type\": \"FactSet\",  \"facts\": [   ' + $scoreJson + '  ]  } ] }'

m365 adaptivecard send --url $webhookUrl --card $card
```

Keywords:

- Microsoft Teams
- Adoption
