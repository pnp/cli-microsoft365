# List app usage in Microsoft Teams

Author: [Arjun Menon](https://twitter.com/arjunumenon)

A sample script which iterates through all the teams in your tenant and lists all apps in each team. This script will be handy if you want to generate a report of available apps in Teams across your tenant.

```powershell tab="PowerShell Core"
$availableTeams = o365 teams team list -o json | ConvertFrom-Json

if ($availableTeams.count -gt 15) {
    $duration =  [math]::Round(($availableTeams.count/60),1);
    Write-Host "There are total of $($availableTeams.count) teams. This probably will take around $duration minutes to finish."
} else {
    Write-Host "There are total of $($availableTeams.count) teams."
}

foreach ($team in $availableTeams) {
    $apps = o365 teams app list -i $team.Id -a    
    Write-Output "All apps in team are given below: $($team.displayName) $($team.id)"
    Write-Output $apps
}
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/
defaultIFS=$IFS
IFS=$'\n'

availableTeams=$(o365 teams team list -o json)

if [[ $(echo $availableTeams | jq length) -gt 15 ]]; 
then
  duration=$(((($(echo $availableTeams | jq length)) + 59) / 60))
  echo "There are total of" $(echo $availableTeams | jq length) "teams. This probably will take around" $duration" minutes to finish."
else
  echo "There are total of" $(echo $availableTeams | jq length) "teams available"
fi

for team in $(echo $availableTeams | jq -c '.[]'); do
    apps=$(o365 teams app list -i $(echo $team | jq ''.id) -a)
    echo "All apps in team are given below: " $(echo $team | jq ''.displayName) " " $(echo $team | jq ''.id)
    echo $apps
done
```

Keywords:

- Microsoft Teams
- Governance