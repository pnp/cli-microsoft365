# aad oauth2grant add

Grant the specified service principal OAuth2 permissions to the specified resource

## Usage

```sh
m365 aad oauth2grant add [options]
```

## Options

`-h, --help`
: output usage information

`-i, --clientId <clientId>`
: `objectId` of the service principal for which permissions should be granted

`-r, --resourceId <resourceId>`
: `objectId` of the AAD application to which permissions should be granted

`-s, --scope <scope>`
: Permissions to grant

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Before you can grant service principal OAuth2 permissions, you need its `objectId`. You can retrieve it using the [aad sp get](../sp/sp-get.md) command.

The resource for which you want to grant permissions is designated using its `objectId`. You can retrieve it using the [aad sp get](../sp/sp-get.md) command, the same way you would retrieve the `objectId` of the service principal.

When granting OAuth2 permissions, you have to specify which permission scopes you want to grant the service principal. You can get the list of available permission scopes either from the resource documentation or from the `appRoles` property when retrieving information about the service principal using the [aad sp get](../sp/sp-get.md) command. Multiple permission scopes can be specified separated by a space.

When granting OAuth2 permissions, the values of the `clientId` and `resourceId` properties form a unique key. If a grant for the same `clientId`-`resourceId` pair already exists, running the `aad oauth2grant add` command will fail with an error. If you want to change permissions on an existing OAuth2 grant use the [aad oauth2grant set](./oauth2grant-set.md) command instead.

## Examples

Grant the service principal _d03a0062-1aa6-43e1-8f49-d73e969c5812_ the _Calendars.Read_ OAuth2 permissions to the _c2af2474-2c95-423a-b0e5-e4895f22f9e9_ resource.

```sh
m365 aad oauth2grant add --clientId d03a0062-1aa6-43e1-8f49-d73e969c5812 --resourceId c2af2474-2c95-423a-b0e5-e4895f22f9e9 --scope Calendars.Read
```

Grant the service principal _d03a0062-1aa6-43e1-8f49-d73e969c5812_ the _Calendars.Read_ and _Mail.Read_ OAuth2 permissions to the _c2af2474-2c95-423a-b0e5-e4895f22f9e9_ resource.

```sh
m365 aad oauth2grant add --clientId d03a0062-1aa6-43e1-8f49-d73e969c5812 --resourceId c2af2474-2c95-423a-b0e5-e4895f22f9e9 --scope "Calendars.Read Mail.Read"
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)