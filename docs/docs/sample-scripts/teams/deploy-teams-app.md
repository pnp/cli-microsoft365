# Deploy Microsoft Teams app from Azure DevOps

Author: [Garry Trinder](https://github.com/garrytrinder)

Installs or updates a Microsoft Teams app from an Azure DevOps pipeline. Deploys the app if it hasn't been deployed yet or updates the existing package if it's been previously deployed.

```powershell tab="PowerShell Core"
m365 login -t password -u $(username) -p $(password)

$apps = m365 teams app list -o json | ConvertFrom-Json
$app = $apps | Where-Object { $_.externalId -eq $env:APPID}
if ($app -eq $null) {
  # install app
  m365 teams app publish -p  $(System.DefaultWorkingDirectory)/teams-app-CI/package/teams-app.zip
} else {
  # update app
  m365 teams app update -i $app.id -p $(System.DefaultWorkingDirectory)/teams-app-CI/package/teams-app.zip
}
```

```bash tab="Bash"
m365 login -t password -u $(username) -p $(password)

app=$(m365 teams app list -o json | jq '.[] | select(.externalId == "'"$APPID"'")')

if [ -z "$app" ]; then
  # install app
  m365 teams app publish -p "$(System.DefaultWorkingDirectory)/teams-app-CI/package/teams-app.zip"
else
  # update app
  appId=$(echo $app | jq '.id')
  m365 teams app update -i $appId -p "$(System.DefaultWorkingDirectory)/teams-app-CI/package/teams-app.zip"
fi
```

Keywords:

- Microsoft Teams
- Azure DevOps
- Continuous Deployment
