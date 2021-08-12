# Cancel all running flow runs for a flow in an environment

Author: [Mohamed Ashiq Faleel](https://ashiqf.com/2021/05/16/cancel-all-your-running-power-automate-flow-runs-using-m365-cli-and-rest-api/)

Do you want to automate the cancellation of running Power Automate flow runs?

Microsoft 365 CLI cmdlets will help you to cancel all the running flow runs.

This script will cancel all running flow runs of a Power Automate flow created in an environment. Pass the Flow environment id and the flow guid as parameter while running the script.

```powershell tab="PowerShell"
$flowEnvironment = $args[0]
$flowGUID = $args[1]
$flowRuns = m365 flow run list --environment $flowEnvironment --flow $flowGUID --output json | ConvertFrom-Json
foreach ($run in $flowRuns) {
  if ($run.status -eq "Running") {
    Write-Output "Run details: " $run
    # Cancel all the running flow runs
    m365 flow run cancel --environment $flowEnvironment --flow $flowGUID --name $run.name --confirm
    Write-Output "Run Cancelled successfully"
  }
}
```
