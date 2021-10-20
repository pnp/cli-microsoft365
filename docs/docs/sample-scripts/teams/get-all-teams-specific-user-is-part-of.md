# Get all the Teams a specific user is part of

Author: [Veronique Lengelle](https://twitter.com/veronicageek)

## Script

```powershell tab="PowerShell Core"
#Variables
$logFile = "<YOUR-FILE-PATH>"
$userToFind = "john.doe@contoso.com"
$results = @()

#Get all the Teams
$allTeams = m365 teams team list -o json | ConvertFrom-Json

#Find the user in Azure AD
$userToFindInAD = m365 aad user get --userName $userToFind -o json | ConvertFrom-Json
$userToFindID = $userToFindInAD.Id

#Loop thru all the Teams
foreach($team in $allTeams){
    $allTeamsUsers = m365 teams user list --teamId $team.Id -o json | ConvertFrom-Json
    
    #Loop through users TARGETING THE USER ID TO MATCH
    foreach ($teamUser in $allTeamsUsers) {
        if ($teamUser.Id -match $userToFindID) {
        
            $results += [pscustomobject]@{
                userName        = $userToFindInAD.UserPrincipalName
                userDisplayName = $userToFindInAD.DisplayName
                userRole        = $teamUser.UserType
                Team            = $team.DisplayName
                ArchivedTeam    = $team.isArchived
                TeamID          = $team.Id
            }
        }
    }    
}
$results | Export-Csv -Path $logFile -NoTypeInformation
```

## Function

```powershell tab="PowerShell Core"
<#
.SYNOPSIS
    Get all the Microsoft Teams team(s) a specific user is part of.
.DESCRIPTION
    Get all the Microsoft Teams team(s) a specific user is part of, and exports the results into a CSV file.
.EXAMPLE
    PS C:\> Get-TeamsUserIsPartOf -UserToFind "john.doe@contoso.com" -logFile "C:\users\$env:USERNAME\Desktop\myFileExport.csv"
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    The export will contain the username (UPN), user display name, user role in that Team, Team display name, Team archive status, and the Team Id.
#>
function Get-TeamsUserIsPartOf {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "User's UPN")]
        [string]$userToFind,
        [Parameter(Mandatory = $true, HelpMessage = "Full path of your .csv file to log the results")]
        [string]$logFile
    )
    $results = @()

    #Get all the Teams
    $allTeams = m365 teams team list -o json | ConvertFrom-Json

    #Find the user in Azure AD
    $userToFindInAD = m365 aad user get --userName $userToFind -o json | ConvertFrom-Json
    $userToFindID = $userToFindInAD.Id

    #Loop thru all the Teams
    foreach($team in $allTeams){
        $allTeamsUsers = m365 teams user list --teamId $team.Id -o json | ConvertFrom-Json
        
        #Loop through users TARGETING THE USER ID TO MATCH
        foreach ($teamUser in $allTeamsUsers) {
            if ($teamUser.Id -match $userToFindID) {
            
                $results += [pscustomobject]@{
                    userName        = $userToFindInAD.UserPrincipalName
                    userDisplayName = $userToFindInAD.DisplayName
                    userRole        = $teamUser.UserType
                    Team            = $team.DisplayName
                    ArchivedTeam    = $team.isArchived
                    TeamID          = $team.Id
                }
            }
        }    
    }
    $results | Export-Csv -Path $logFile -NoTypeInformation
}
```

Keywords

- Microsoft Teams
- Governance
