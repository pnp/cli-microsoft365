# Persisting connection information

After logging in to Office 365, the Office 365 CLI will persist the information about the connection until you explicitly log out from Office 365.

## Why is persisting connection information important

Persisting connection information is important for two reasons.

### Convenience

First of all it's more convenient to use the Office 365 CLI. If you're using it often to manage a specific tenant, you can connect once and the CLI will remember your credentials. The next time you start the CLI, you can directly start managing your tenant without having to authenticate first.

### Building scripts

When working with Office 365 CLI, each command is executed in an isolated context and has no access to the memory of any command executed before. So unless you would store the connection information in a variable and explicitly pass it to each command, the CLI would be unable to log in to your tenant. As you can imagine, working with the CLI in this way would be tedious and inconvenient.

By persisting the connection information the Office 365 CLI can be used to build scripts, for example:

_Deploy all apps that are not yet deployed in the tenant app catalog:_

```sh
# get all apps available in the tenant app catalog
apps=$(o365 spo app list -o json)

# get IDs of all apps that are not deployed
notDeployedAppsIds=($(echo $apps | jq -r '.[] | select(.Deployed == false) | {ID} | .[]'))

# deploy all not deployed apps
for appId in $notDeployedAppsIds; do
  o365 spo app deploy -i $appId
done
```

First, you use the Office 365 CLI to get the list of all apps from the tenant app catalog using the `spo app list` command. You set the output type to JSON and store it in a shell variable `apps`. Next, you parse the JSON string using [jq](https://stedolan.github.io/jq/) and get IDs of apps that are not deployed. Finally, for each ID you run the `spo app deploy` Office 365 CLI command passing the ID as a command argument. Notice, that in the script, both `spo` commands are run as separate commands directly in the shell. In both cases, the shell starts the CLI, executes the specified command and closes the CLI removing all of its resources from memory. Scripts, like the one mentioned above can work, because the Office 365 CLI persists its connection information.

## Persisting connection information in Office 365 CLI

When you log in to Office 365 in the Office 365 CLI, the CLI will persist the information about the connection for future reuse. For the established connection, the Office 365 CLI persists the refresh token as well as all access tokens obtained when using the different CLI commands.

Depending on the Office 365 CLI commands you have used, the CLI might persist some additional information. For example, when using commands that interact with SharePoint Online, the CLI will store the URL of your SharePoint Online tenant as well as its ID.

The Office 365 CLI stores its connection information in a JSON file located in the home directory of the current user, on MacOS and Linux, this is `~/.o365cli-tokens.json` and on Windows, this is `<root>\Users\<username>\.o365cli-tokens.json`. The contents of this file are not encrypted.

## Removing persisted connection information

Office 365 CLI persists its connection information until you either explicitly log out from the particular service or manually remove the stored credentials.

To check if you are logged in to Office 365 in the Office 365 CLI, run the `status` command. If you are logged in, the command will return the name of the user account or AAD application used to log in. If you are not connected, the command will return `false`.

To log out from Office, run the `logout` command. Running this command will remove all previously stored connection data from your machine.