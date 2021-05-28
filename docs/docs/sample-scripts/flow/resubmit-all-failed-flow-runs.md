# Resubmit all failed flow runs for a flow in an environment

Author: [Mohamed Ashiq Faleel](https://ashiqf.com/2021/05/09/resubmit-your-failed-power-automate-flow-runs-automatically-using-m365-cli-and-rest-api/)

Have you ever been forced to resubmit lot of failed Power Automate flow runs manually?

Microsoft 365 CLI cmdlets to the rescue, it will help you resubmit the flow runs automatically.

This script will resubmit all failed flow runs of a Power Automate flow created in an environment. Pass the Flow environment id and the flow guid as parameter while running the script.

```powershell tab="PowerShell Core"
$flowEnvironment = $args[0]
$flowGUID = $args[1]
$flowRuns = m365 flow run list --environment $flowEnvironment --flow $flowGUID --output json | ConvertFrom-Json
foreach ($run in $flowRuns) {
  if ($run.status -eq "Failed") {
    Write-Output "Run details: " $run
    #Resubmit all the failed flows
    m365 flow run resubmit --environment $flowEnvironment --flow $flowGUID --name $run.name --confirm
    Write-Output "Run resubmitted successfully"
  }
}
```
