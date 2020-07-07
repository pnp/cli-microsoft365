# Office 365 CLI output mode

Office 365 CLI commands can present their output either as plain-text or as JSON. Following is information on these two output modes along with information when to use which.

## Choose the command output mode

All commands in Office 365 CLI can present their output as plain-text or as JSON. By default, all commands use the text output mode, but by setting the `--output`, or `-o` for short, option to `json`, you can change the output mode for that command to JSON.

## Text output mode

By default, all commands in Office 365 CLI present their output as plain-text. Depending on the command output, the value is presented as-is or formatted for readability.

### Simple values

If the command output is a simple value, such as a number, boolean or a string, the value is returned as-is. For example, if the Office 365 Public CDN is enabled on the currently connected tenant, executing the `spo cdn get` command, will return `true`:

```sh
$ o365 spo cdn get
true
```

### Objects

If the command returns information about an object such as a site, list or an app, that contains a number of properties, the output in text mode is formatted as key-value pairs. For example, getting information about a specific app, will return output similar to:

```sh
$ o365 spo app get -i e6362993-d4fd-4c5a-8254-fd095a7291ad
AppCatalogVersion     : 1.0.0.0
CanUpgrade            : false
CurrentVersionDeployed: false
Deployed              : false
ID                    : e6362993-d4fd-4c5a-8254-fd095a7291ad
InstalledVersion      :
IsClientSideSolution  : true
Title                 : spfx-140-online-client-side-solution
```

### Arrays

If the command returns information about multiple objects, the output is formatted as a table. For example, getting the list of available app, will return output similar to:

```sh
$ o365 spo app list
Title                                 ID                                    Deployed  AppCatalogVersion
------------------------------------  ------------------------------------  --------  -----------------
spfx-140-online-client-side-solution  e6362993-d4fd-4c5a-8254-fd095a7291ad  false     1.0.0.0
spfx-134-client-side-solution         5ae74650-b00b-46a9-925f-9c9bd70a0cb6  false     1.0.0.0
```

If only one app is returned, it will be displayed as key-value pairs:

```sh
$ o365 spo app list
AppCatalogVersion: 1.0.0.0
Deployed         : false
ID               : e6362993-d4fd-4c5a-8254-fd095a7291ad
Title            : spfx-140-online-client-side-solution
```

## JSON output mode

Optionally, you can have all Office 365 CLI commands return their output as JSON strings. This is invaluable when building scripts using the CLI, where the output of one command, has to be processed by another command.

### Simple values

Simple values in JSON output, are returned as-is. For example, if the Office 365 Public CDN is enabled on the currently connected tenant, executing the `spo cdn get` command, will return `true`:

```sh
$ o365 spo cdn get -o json
true
```

### Objects

If the command returns an object, that object will be formatted as a JSON string. For example, getting information about a specific app, will return output similar to:

```sh
$ m365 spo app get -i e6362993-d4fd-4c5a-8254-fd095a7291ad
{
  "AppCatalogVersion": "1.0.0.0",
  "CanUpgrade": false,
  "CurrentVersionDeployed": false,
  "Deployed": false,
  "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
  "InstalledVersion": "",
  "IsClientSideSolution": true,
  "Title": "spfx-140-online-client-side-solution"
}
```

### Arrays

If the command returns information about multiple objects, the command will return a JSON array with each array item representing one object. For example, getting the list of available app, will return output similar to:

```sh
$ m365 spo app list -o json
[
  {
    "AppCatalogVersion": "1.0.0.0",
    "CanUpgrade": false,
    "CurrentVersionDeployed": false,
    "Deployed": false,
    "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
    "InstalledVersion": "",
    "IsClientSideSolution": true,
    "Title": "spfx-140-online-client-side-solution"
  },
  {
    "AppCatalogVersion": "1.0.0.0",
    "CanUpgrade": false,
    "CurrentVersionDeployed": false,
    "Deployed": false,
    "ID": "5ae74650-b00b-46a9-925f-9c9bd70a0cb6",
    "InstalledVersion": "",
    "IsClientSideSolution": true,
    "Title": "spfx-134-client-side-solution"
  }
]
```

Even if the array contains only one item, for consistency it will be returned as a one-element JSON array:

```sh
$ m365 spo app list -o json
[
  {
    "AppCatalogVersion": "1.0.0.0",
    "CanUpgrade": false,
    "CurrentVersionDeployed": false,
    "Deployed": false,
    "ID": "e6362993-d4fd-4c5a-8254-fd095a7291ad",
    "InstalledVersion": "",
    "IsClientSideSolution": true,
    "Title": "spfx-140-online-client-side-solution"
  }
]
```

!!! tip
    Some `list` commands return different output in text and JSON mode. For readability, in the text mode they only include a few properties, so that the output can be formatted as a table and will fit on the screen. In JSON mode however, they will include all available properties so that it's possible to process the full set of information about the particular object. For more details, refer to the help of the particular command.

### Verbose and debug output in JSON mode

When executing commands in JSON output mode with the `--verbose` or `--debug` flag, the verbose and debug logging statements will be also formatted as JSON and will be added to the output. When processing the command output, you would have to determine yourself which of the returned JSON objects represents the actual command result and which are additional verbose and debug logging statements.

## Processing command output with JMESPath

Office 365 CLI supports filtering, sorting and querying data returned by its commands using [JMESPath](http://jmespath.org/) queries. Queries can be specified using the `--query` option on each command and are applied just before the data retrieved by the command is sent to the console. While you can apply JMESPath queries in all output modes, they are the most powerful in combination with JSON output where the data is unfiltered and the complete data set is sent to output.

For example, you can retrieve the list of all SharePoint site collections in your tenant, by executing:

```sh
$ o365 spo site list
Title                                Url
-----------------------------------  -------------------------------------------------------------------------
Digital Initiative Public Relations  https://contoso.sharepoint.com/sites/DigitalInitiativePublicRelations
Leadership Team                      https://contoso.sharepoint.com/sites/leadership
Mark 8 Project Team                  https://contoso.sharepoint.com/sites/Mark8ProjectTeam
Operations                           https://contoso.sharepoint.com/sites/operations
Retail                               https://contoso.sharepoint.com/sites/Retail
Sales and Marketing                  https://contoso.sharepoint.com/sites/SalesAndMarketing
```

To retrieve information only about sites matching a specific title or URL, you could execute:

```sh
$ o365 spo site list --query "[?Title == 'Retail']"
Title: Retail
Url  : https://contoso.sharepoint.com/sites/Retail
```

!!! important
    Please note, that because you haven't specified the output mode, the command uses its default text output mode and you can only query the `Title` and `Url` properties of each site, because only these properties are returned by the command in the text output mode.

If you wanted to find site collections using a specific site template, you would switch to the JSON output mode, to get access to all retrieved information, and apply the query to the returned data:

```sh
$ o365 spo site list --output json --query "[?Template == 'GROUP#0'].{Title: Title, Url: Url}"
[{"Title":"Mark 8 Project Team","Url":"https://contoso.sharepoint.com/sites/Mark8ProjectTeam"},{"Title":"Operations","Url":"https://contoso.sharepoint.com/sites/operations"},{"Title":"Digital Initiative Public Relations","Url":"https://contoso.sharepoint.com/sites/DigitalInitiativePublicRelations"},{"Title":"Retail","Url":"https://contoso.sharepoint.com/sites/Retail"},{"Title":"Leadership Team","Url":"https://contoso.sharepoint.com/sites/leadership"},{"Title":"Sales and Marketing","Url":"https://contoso.sharepoint.com/sites/SalesAndMarketing"}]
```

To make the output more readable, you could pass it to a JSON processor such as [jq](https://stedolan.github.io/jq/):

```sh
$ o365 spo site list --output json --query "[?Template == 'GROUP#0'].{Title: Title, Url: Url}" | jq
[
  {
    "Title": "Mark 8 Project Team",
    "Url": "https://contoso.sharepoint.com/sites/Mark8ProjectTeam"
  },
  {
    "Title": "Operations",
    "Url": "https://contoso.sharepoint.com/sites/operations"
  },
  {
    "Title": "Digital Initiative Public Relations",
    "Url": "https://contoso.sharepoint.com/sites/DigitalInitiativePublicRelations"
  },
  {
    "Title": "Retail",
    "Url": "https://contoso.sharepoint.com/sites/Retail"
  },
  {
    "Title": "Leadership Team",
    "Url": "https://contoso.sharepoint.com/sites/leadership"
  },
  {
    "Title": "Sales and Marketing",
    "Url": "https://contoso.sharepoint.com/sites/SalesAndMarketing"
  }
]
```

## When to use which output mode

Generally, you will use the default text output when interacting with the CLI yourself. When building scripts using the Office 365 CLI, you will use the JSON output mode, because processing JSON strings is much more convenient and reliable than processing plain-text output.

## Sample script

Using the JSON output mode allows you to build scripts using the Office 365 CLI. The CLI works on any platform, but as there is no common way to work with objects and command output on all platforms and shells, we chose JSON as the format to serialize objects in the Office 365 CLI.

Following, is a sample script, that you could build using the Office 365 CLI in Bash:

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

_First, you use the Office 365 CLI to get the list of all apps from the tenant app catalog using the [spo app list](../cmd/spo/app/app-list.md) command. You set the output type to JSON and store it in a shell variable `apps`. Next, you parse the JSON string using [jq](https://stedolan.github.io/jq/) and get IDs of apps that are not deployed. Finally, for each ID you run the [spo app deploy](../cmd/spo/app/app-deploy.md) Office 365 CLI command passing the ID as a command argument. Notice, that in the script, both `spo` commands are prepended with `o365` and executed as separate commands directly in the shell._

The same could be accomplished in PowerShell as well:

```powershell
# get all apps available in the tenant app catalog
$apps = o365 spo app list -o json | ConvertFrom-Json

# get all apps that are not yet deployed and deploy them
$apps | ? Deployed -eq $false | % { o365 spo app deploy -i $_.ID }
```

Because PowerShell offers native support for working with JSON strings and objects, the same script written in PowerShell is simpler than the one in Bash. At the end of the day it's up to you to choose if you want to use the Office 365 CLI in Bash, PowerShell or some other shell. Both Bash and PowerShell are available on multiple platforms, and if you have a team using different platforms, writing scripts using Office 365 CLI in Bash or PowerShell will let everyone in your team use them.