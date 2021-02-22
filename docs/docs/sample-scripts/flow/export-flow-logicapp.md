# Export a single Flow to a Logic App

Author: [Albert-Jan Schot](https://www.cloudappie.nl/cli-m365-exportflow/)

!!!

Want to upgrade one of your Flows to a Logic App? Missing the option in the UI? Or just looking at an easy way to do it programmatically?

By combining the CLI for Microsoft 365 and PowerShell we can make this task easy and repeatable.

This script will export the flow *Your sample test flow*, make sure to pass the correct name in the script, and your Flow will be exported right away.
!!!

```powershell tab="PowerShell Core"
Write-Output "Getting environment info..."
$environmentId = $(m365 flow environment list --query "[?displayName == '(default)']" -o json | ConvertFrom-Json).Name
$flowId = $(m365 flow list --environment $environmentId --query "[?displayName == 'Your sample test flow']" -o json | ConvertFrom-Json)[0].Name

Write-Output "Getting Flow info..."
m365 flow export --environment $environmentId --id $flowId -f 'json'

Write-Output "Complete"
```

```bash tab="Bash"
#!/bin/bash
ENV_NAME=m365 flow environment list --query '[?contains(displayName,`default`)] .name'
FLOW_NAME=m365 flow list --environment $environmentId --query '[?displayName == `Your sample test flow`] .name'
echo "Exporting your flow"
m365 flow export --environment $ENV_NAME --id $FLOW_NAME -f 'json'
```
