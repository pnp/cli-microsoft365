# aad approleassignment add

Deletes an app role assignment for the specified Azure AD Application Registration

## Usage

```sh
m365 aad approleassignment remove [options]
```

## Options

`-h, --help`
: output usage information

`--appId [appId]`
: Application appId also known as clientId of the App Registration for which the configured scopes (app roles) should be deleted

`--objectId [objectId]`
: Application objectId of the App Registration for which the configured scopes (app roles) should be deleted

`--displayName [displayName]`
: Application name of the App Registration for which the configured scopes (app roles) should be deleted

`-r, --resource <resource>`
: Service principal name, appId or objectId that has the scopes (roles) ex. `SharePoint`

`-s, --scope <scope>`
: Permissions known also as scopes and roles to be deleted from the application. If multiple permissions have to be deleted, they have to be comma separated ex. `Sites.Read.All`,`Sites.ReadWrite.All`

`--confirm`
: Don't prompt for confirming removing the all role assignment

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

This command requires tenant administrator permissions.

Specify either the `appId`, `objectId` or `displayName` but not both. If you specify more than one option value, the command will fail with an error.

Autocomplete values for the `resource` option do not mean allowed values. The autocomplete will just suggest some known names, but that doesn't restrict you to use name of your own custom application or other application within your tenant.

This command can also be used to assign permissions to system- or user-assigned managed identity.

## Examples

Deletes SharePoint _Sites.Read.All_ application permissions from Azure AD application with app id _57907bf8-73fa-43a6-89a5-1f603e29e451_

```sh
m365 aad approleassignment remove --appId "57907bf8-73fa-43a6-89a5-1f603e29e451" --resource "SharePoint" --scope "Sites.Read.All"
```

Deletes multiple Microsoft Graph application permissions from an Azure AD application with name _MyAppName_

```sh
m365 aad approleassignment remove --displayName "MyAppName" --resource "Microsoft Graph" --scope "Mail.Read,Mail.Send"
```

Deletes Microsoft Graph _Mail.Read_ application permissions from a system managed identity app with objectId _57907bf8-73fa-43a6-89a5-1f603e29e451_

```sh
m365 aad approleassignment remove --objectId "57907bf8-73fa-43a6-89a5-1f603e29e451" --resource "Microsoft Graph" --scope "Mail.Read"
```

## More information

- Microsoft Graph permissions reference: [https://docs.microsoft.com/en-us/graph/permissions-reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
