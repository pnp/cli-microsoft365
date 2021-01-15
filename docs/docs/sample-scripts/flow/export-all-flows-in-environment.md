# Export all Flows in environment

Author: [Garry Trinder](https://garrytrinder.github.io/2021/01/export-all-flows-from-environment-cli-microsoft365)

When was the last time you backed up all the Flows in your environment?

By combining the CLI for Microsoft 365 and PowerShell we can make this task easy and repeatable.

This script will get all Flows in your default environment and export them as both a ZIP file for importing back into Power Automate and as a JSON file for importing into Azure as an Azure Logic App.

```powershell tab="PowerShell Core"
Write-Output "Getting environment info..."
$environment = m365 flow environment list --query '[?contains(displayName,`default`)] .name'

Write-Output "Getting Flows info..."
$flows = m365 flow list --environment $environment --asAdmin --output json | ConvertFrom-JSON

Write-Output "Found $($flows.Count) Flows to export..."

$flows | ForEach-Object {
    Write-Output "Exporting as ZIP & JSON... $($_.displayName)"
    $filename = $_.displayName.Replace(" ","")
    $timestamp = Get-Date -Format "yyyymmddhhmmss"
    $exportPath = "$($filename)_$($timestamp)"
    $flowId = $_.Name
    
    m365 flow export --id $flowId --environment $environment --packageDisplayName $_.displayName --path "$exportPath.zip"
    m365 flow export --id $flowId --environment $environment --format json --path "$exportPath.json"
}

Write-Output "Complete"
```