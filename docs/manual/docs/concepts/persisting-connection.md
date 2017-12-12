# Persisting connection information

After connecting to an Office 365 service, like SharePoint Online, the Office 365 CLI will persist the information about the connection until you explicitly disconnect from the specific service.

## Why is persisting connection information important

Persisting connection information is important for two reasons.

### Convenience

First of all it's more convenient to use the Office 365 CLI. If you're using it often to manage a specific tenant, you can connect once and the CLI will remember your credentials. The next time you start the CLI, you can directly start managing your tenant without having to authenticate first.

### Building scripts

Additionally, it makes it possible for you to write scripts that automate management of your tenant. When the Office 365 CLI is run in immersive (interactive) mode, the connection information is persisted in memory and is available to all commands run in the CLI command prompt. Unfortunately, one limitation of the immersive mode is that you can only run one command at a time and can't pass the output of one command into another.

When running in the non-immersive (non-interactive) mode, each command is executed in an isolated context and has no access to the memory of any command executed before. So unless you would store the connection information in a variable and explicitly pass it to each command, the CLI would be unable to connect to your tenant. As you can imagine, working with the CLI in this way would be tedious and inconvenient.

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

When you connect in the Office 365 CLI to an Office 365 service, such as SharePoint Online, the CLI will persist the information about the connection for future reuse. For each established connection, the Office 365 CLI persists the following information:

- service name, eg. `SPO`
- Azure AD resource name, eg `https://contoso.sharepoint.com`
- refresh token
- access token
- access token expiration timestamp

Depending on the Office 365 service to which you connect, the Office 365 CLI might persist some additional information. For example, when connecting to SharePoint Online tenant admin site using the `spo connect` command, the CLI will store the tenant ID. If you were initially connected to the tenant admin site, but also performed operations on other site collections (like retrieving the list of apps installed in the specific site), the CLI will store access token for regular SharePoint sites as well.

Where the connection information is persisted, depends on the operating system that you are using.

### macOS

On the macOS, the Office 365 CLI persists its connection information in the system Keychain. For each connected Office 365 service (such as `SPO`) it adds a generic credential. You can view what information is stored by opening **Keychain Access** and searching for `Office 365 CLI`.

### Windows

On Windows, the Office 365 CLI persists its connection information in the Windows Credential Manager. To view the persisted credentials, from the **Control Panel**, navigate to **User Accounts** and from the **Credential Manager** section open **Manage Windows Credentials**. Any credentials stored by the Office 365 CLI will be listed in the **Generic Credentials** section named as `[service]--x-y`, for example `SPO--1-3`. Because there is a limit how long a password stored in the Windows Credential Manager can be, connection information stored by the Office 365 CLI will often be split over multiple chunks, where the last two number in the chunk specify the number of chunk and the total number of chunks.

### Linux

On Linux, the Office 365 CLI stores its connection information in a JSON file located at `~/.o365cli-tokens.json`. The contents of this file are not encrypted. The primary use case for supporting Linux operating system is to use the Office 365 CLI in Docker containers, where the tokens file is persisted in the container as long as the container is running. When the container is closed and removed, the file is removed as well. When you would start the container again, you would have to connect to Office 365 first, before you could use the Office 365 CLI.

## Removing persisted connection information

Office 365 CLI persists its connection information until you either explicitly disconnect from the particular service or manually remove the stored credentials.

To check if you are connected to a particular Office 365 service in the Office 365 CLI, run the corresponding status command, for example `o365 spo status`. If you are connected, the command will return the URL of the site to which you are connected. If you are not connected, the command will return `false`.

To disconnect from the specific Office 365 service, run the corresponding Office 365 CLI, for example, to disconnect from SharePoint Online and remove all persisted connection information, run `o365 spo disconnect`.