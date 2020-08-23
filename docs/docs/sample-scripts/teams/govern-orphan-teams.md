# Govern orphaned Microsoft Teams

Author: [Matti Paukkonen](https://mattipaukkonen.com/2019/10/09/govern-orphaned-teams-with-powershell/)

Every team needs an owner, at least one. Common best practice is that you should have at least two users in owner role. Teams is not allowing the last owner to leave the team, but there might occasions when last owner is removed, example when people are leaving the organization and account gets deleted. This script finds those teams that no longer have an owner.

```powershell tab="PowerShell Core"
$availableTeams = m365 teams team list -o json | ConvertFrom-Json
$teams = @()
foreach ($team in $availableTeams) {

    Write-host "Handling team: " -NoNewline -ForegroundColor Yellow
    Write-host $team.DisplayName -ForegroundColor Yellow
    $group = m365 aad o365group get --id $team.id -o json | ConvertFrom-Json
    $users = m365 teams user list --teamId $team.id -o json | ConvertFrom-Json
    $owners = @($users | Where-Object { $_.userType -eq "Owner" })
    $members = @($users | Where-Object { $_.userType -eq "Member" }).Length
    $guests = @($users | Where-Object { $_.userType -eq "Guest" }).Length

    $teamObject = New-Object -TypeName PSObject
    $teamObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $team.displayName
    $teamObject | Add-Member -MemberType NoteProperty -Name Alias -Value $group.mailNickName
    $teamObject | Add-Member -MemberType NoteProperty -Name "Number of Owners" -Value $owners.Length
    $teamObject | Add-Member -MemberType NoteProperty -Name "Number of Members" -Value $members
    $teamObject | Add-Member -MemberType NoteProperty -Name "Number of Guests" -Value $guests
    if ($owners.Count -eq 1) {
        $teamObject | Add-Member -MemberType NoteProperty -Name "Owner" -Value $owners[0].displayName
    }

    write-host " ...Done" -ForegroundColor Green
    $teams += $teamObject
}

$teams | Format-Table -AutoSize
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

defaultIFS=$IFS
IFS=$'\n'

availableTeams=`m365 teams team list -o json`
teams=()

for team in `echo $availableTeams | jq -c '.[]'`; do

    displayName=`echo $team | jq '.displayName'`
    echo "Handling team: ${displayName}"

    teamId=`echo $team | jq '.id'`
    group=`m365 aad o365group get --id ${teamId} -o json`
    users=`m365 teams user list --teamId ${teamId} -o json`

    groupId=`echo $team | jq '.id'`
    alias=`echo $group | jq '.mailNickName'`

    owner=`echo $users | jq -c 'map(select(.userType == "Owner")) | .[0]? | .displayName'`
    ownercount=`echo $users | jq -c 'map(select(.userType == "Owner")) | length'`
    membercount=`echo $users | jq -c 'map(select(.userType == "Member")) | length'`
    guestcount=`echo $users | jq -c 'map(select(.userType == "Guest")) | length'`

    teamObject=$(jq -n -c \
        --arg dn "${displayName}" \
        --arg id "${groupId}" \
        --arg al "${alias}" \
        --arg oc "${ownercount}" \
        --arg mc "${membercount}" \
        --arg gc "${guestcount}" \
        --arg ow "${owner}" \
        '{DisplayName: $dn, GroupID: $id, Alias: $al, OwnerCount: $oc, MemberCount: $mc, GuestCount: $gc, Owner: $ow}')

    echo "...Done"
    teams+=($teamObject)
done

echo ${teams[@]} | jq -csr '(.[0] |keys_unsorted | @tsv), (.[]|.|map(.) |@tsv)' | column -s$'\t' -t

IFS=defaultIFS

exit 1
```

Keywords:

- Microsoft Teams
- Governance
