# Bulk add members to Microsoft Teams team from CSV file

Inspired by: [Rakesh Pandey](https://www.flexmind.co/blog/how-to-add-bulk-users-from-csv-file-to-ms-teams-using-powershell/)

```powershell tab="PowerShell"
<#
.SYNOPSIS
    Add users to a Microsoft 365 group linked to Teams.
.DESCRIPTION
    This script will add users in existing Teams contained in your .csv file.
.EXAMPLE
    PS C:\> .\add-users-teams.ps1
    This script will add users in existing Microsoft Teams teams from your .csv file
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    Your .csv file must contain headers called UPN, teamName, teamId, and role. If you change those headers then make sure to amend the script.
#>

#Check if connected to M365
$m365Status = m365 status
if ($m365Status -eq "Logged Out") {
  m365 login
}
    
#Import csv
$usersCsvFile = Import-Csv -Path "<YOUR_CSVFile_PATH_HERE.csv>"

#Add users to the Team
foreach ($row in $usersCsvFile) {
  Write-Host "Adding $($row.UPN) to the $($row.teamName) Team" -ForegroundColor Magenta
  m365 aad o365group user add --groupId $row.teamId --userName $($row.UPN) --role $($row.role)
}
```

Keywords:

- Teams
- Provisioning
