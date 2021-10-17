# Remove Wiki tab in a Microsoft Teams channel

Inspired by: [Garry Trinder](https://gist.github.com/garrytrinder/4df2aeaf9dd66c4375308874eb7def63) and [Laura Kokkarinen](https://laurakokkarinen.com/deleting-the-treacherous-wiki-tab-as-a-part-of-your-teams-provisioning-process/)

Removes the wiki tab of a Microsoft Teams Team's channel.

=== "PowerShell"

    ```powershell
    $groupMailNickname = "Architecture"
    $channelName = "General"

    $groups = m365 aad o365group list --query "[?mailNickname=='$groupMailNickname']" -o json | ConvertFrom-Json
    if ($null -eq $groups) { Write-Error "A team with the mailNickname $groupMailNickname was not found" }
    else {
      $channels = m365 teams channel list --teamId $groups[0].id --query "[?displayName=='$channelName']" -o json | ConvertFrom-Json
      if ($null -eq $channels) { Write-Error "A channel with the name $channelName was not found in the team" }
      else {
        $tabs = m365 teams tab list --teamId $groups[0].id --channelId $channels[0].id --query "[?teamsApp.id=='com.microsoft.teamspace.tab.wiki']" -o json | ConvertFrom-Json
        if ($null -eq $tabs) { Write-Error "A Wiki tab was not found in the channel" }
        else {
          write-host "Removing wiki tab for the channel.." -ForegroundColor Green 
          m365 teams tab remove --teamId $groups[0].id --channelId $channels[0].id --tabId $tabs[0].id --confirm
          write-host " ...Done" -ForegroundColor Green 
        }
      }
    }
    ```

=== "Bash"

    ```bash
    #!/bin/bash

    # requires jq: https://stedolan.github.io/jq/

    groupMailNickname="Architecture"
    channelName="Channel"
    wikiTabId="com.microsoft.teamspace.tab.wiki"

    groups=$(m365 aad o365group list -o json | jq '.[] | select(.mailNickname == "'"$groupMailNickname"'")')
    if [ -z "$groups" ]; then
      echo "A team with the mailNickname $groupMailNickname was not found"
    else
      teamId=$(echo $groups | jq '.id')
      channels=$(m365 teams channel list --teamId $teamId -o json | jq '.[] | select(.displayName == "'"$channelName"'")')
      
      if [ -z "$channels" ]; then
        echo "A channel with the name $channelName was not found in the team"
      else
        channelId=$(echo $channels | jq '.id')
        tabs=$(m365 teams tab list --teamId $teamId --channelId $channelId -o json | jq '.[] | select(.teamsApp.id == "'"$wikiTabId"'")')

        if [ -z "$tabs" ]; then
          echo "A Wiki tab was not found in the channel"
        else
          tabId=$(echo $tabs | jq '.id')
          echo "Removing wiki tab for the channel.."
          m365 teams tab remove --teamId $teamId --channelId $channelId --tabId $tabId --confirm
          echo "...Done"
        fi
      fi
    fi
    ```

Keywords:

- Microsoft Teams
