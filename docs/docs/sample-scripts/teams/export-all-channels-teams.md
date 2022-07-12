# Export all channels in Microsoft Teams teams in the tenant

Author: [Sudharsan Kesavanarayanan](https://twitter.com/sudharsank)

Export all the channels from Microsoft Team in a CSV.

=== "PowerShell"

    ```powershell
    function Get-Channels(
        [Parameter(Mandatory = $false)][string] $teamID,
        [Parameter(Mandatory = $false)][string] $teamName 
    ) {
        if(!$teamID -and !$teamName) {
            Write-Error "Either 'teamID' or 'teamName' is required"
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
            Write-Host "Exporting file to $fileExportPath.."
            $results | Export-Csv -Path "Channels.csv" -NoTypeInformation
            Write-Host "Completed."
        } else {
            Write-Information "No channels found!"
        }
    }
    
    Write-Host "Ensure logged in"
    $m365Status = m365 status --output text
    if ($m365Status -eq "Logged Out") {
        Write-Host "Logging in the User!"
        m365 login --authType browser
    }
    
    Get-Channels -teamName "<Team Name>"
    Get-Channels -teamID "<Team ID>"
    ```

Keywords:

- Microsoft Teams
- Governance
