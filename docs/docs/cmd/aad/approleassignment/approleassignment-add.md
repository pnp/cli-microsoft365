# aad approleassignment add

Adds service principal permissions also known as scopes and app role assignments for specified Azure AD application registration

## Usage

```sh
m365 aad approleassignment add [options]
```

## Options

`-h, --help`
: output usage information

`--appId [appId]`
: Application appId also known as clientId of the App Registration to which the configured scopes (app roles) should be applied

`--objectId [objectId]`
: Application objectId of the App Registration to which the configured scopes (app roles) should be applied

`--displayName [displayName]`
: Application name of the App Registration to which the configured scopes (app roles) should be applied

`-r, --resource <resource>`
: Service principal name, appId or objectId that has the scopes (roles) ex. `SharePoint`.

`-s, --scope <scope>`
: Permissions known also as scopes and roles to grant the application with. If multiple permissions have to be granted, they have to be comma separated ex. `Sites.Read.All,Sites.ReadWrite.all`

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

This command can also be used to assign permissions to system or user-assigned managed identity.

## Examples

Adds SharePoint _Sites.Read.All_ application permissions to Azure AD application with app id _57907bf8-73fa-43a6-89a5-1f603e29e451_

```sh
m365 aad approleassignment add --appId "57907bf8-73fa-43a6-89a5-1f603e29e451" --resource "SharePoint" --scope "Sites.Read.All"
```

Adds multiple Microsoft Graph application permissions to an Azure AD application with name _MyAppName_

```sh
m365 aad approleassignment add --displayName "MyAppName" --resource "Microsoft Graph" --scope "Mail.Read,Mail.Send"
```

Adds Microsoft Graph _Mail.Read_ application permissions to a system managed identity app with objectId _57907bf8-73fa-43a6-89a5-1f603e29e451_

```sh
m365 aad approleassignment add --objectId "57907bf8-73fa-43a6-89a5-1f603e29e451" --resource "Microsoft Graph" --scope "Mail.Read"
```

## More information

- Microsoft Graph permissions reference: [https://docs.microsoft.com/en-us/graph/permissions-reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
