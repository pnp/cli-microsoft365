# aad oauth2grant list

Lists OAuth2 permission grants for the specified service principal

## Usage

```sh
m365 aad oauth2grant list [options]
```

## Options

`-i, --clientId <clientId>`
: objectId of the service principal for which the configured OAuth2 permission grants should be retrieved

--8<-- "docs/cmd/_global.md"

## Remarks

In order to list existing OAuth2 permissions granted to a service principal, you need its `objectId`. You can retrieve it using the [aad sp get](../sp/sp-get.md) command.

When using the text output type (default), the command lists only the values of the `objectId`, `resourceId` and `scope` properties of the OAuth grant. When setting the output type to JSON, all available properties are included in the command output.

## Examples

List OAuth2 permissions granted to service principal with `objectId` _b2307a39-e878-458b-bc90-03bc578531d6_.

```sh
m365 aad oauth2grant list --clientId b2307a39-e878-458b-bc90-03bc578531d6
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)
- List oauth2PermissionGrants: [https://docs.microsoft.com/en-us/graph/api/oauth2permissiongrant-list?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/oauth2permissiongrant-list?view=graph-rest-1.0)