# List all channels in Microsoft Teams team in the tenant

List all channels in Microsoft Teams team in the tenant and exports the results in a CSV.

```powershell tab="PowerShell"
function Get-Channels(
    [Parameter(Mandatory = $false)][string] $teamID,
    [Parameter(Mandatory = $false)][string] $teamName 
) {
    if(!$teamID -and !$teamName) {
        Write-Error "Either 'Team ID' or 'Team Name' is required"
        return
    }
    $channels = $null
    if($teamID) {
        $channels = m365 teams channel list --teamId $teamID -o 'json' | ConvertFrom-Json
    } 
    if($teamName) {
        $channels = m365 teams channel list --teamName $teamName -o 'json' | ConvertFrom-Json
    }
    Write-Output $channels.length
    if($channels.length -gt 0) {
        $results = @()
        foreach($channel in $channels) {
            $results += [pscustomobject][ordered]@{
                ID = $channel.id
                "Display Name" = $channel.displayName
                Description = $channel.description
                Email = $channel.email
                WebURL = $channel.weburl
            }
        }
        Write-Host "Exporting the results.."
        $results | Export-Csv -Path "Channels.csv" -NoTypeInformation
        Write-Host "Completed."
    } else {
        Write-Information "No channesl found!"
    }
}

Write-Host "Ensure logged in"
$m365Status = m365 status
if ($m365Status -eq "Logged Out") {
    Write-Host "Logging in the User!"
    m365 login --authType browser
}
```

Keywords:

- Microsoft Teams
- Governance
