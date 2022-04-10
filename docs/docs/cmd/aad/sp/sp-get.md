# aad sp get

Gets information about the specific service principal

## Usage

```sh
m365 aad sp get [options]
```

## Options

`-i, --appId [appId]`
: ID of the application for which the service principal should be retrieved

`-n, --displayName [displayName]`
: Display name of the application for which the service principal should be retrieved

`--objectId [objectId]`
: ObjectId of the application for which the service principal should be retrieved

--8<-- "docs/cmd/_global.md"

## Remarks

Specify either the `appId`, `objectId` or `displayName`. If you specify more than one option value, the command will fail with an error.

## Examples

Return details about the service principal with appId _b2307a39-e878-458b-bc90-03bc578531d6_.

```sh
m365 aad sp get --appId b2307a39-e878-458b-bc90-03bc578531d6
```

Return details about the _Microsoft Graph_ service principal.

```sh
m365 aad sp get --displayName "Microsoft Graph"
```

Return details about the service principal with ObjectId _b2307a39-e878-458b-bc90-03bc578531dd_.

```sh
m365 aad sp get --objectId b2307a39-e878-458b-bc90-03bc578531dd
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)
- Get servicePrincipal: [https://docs.microsoft.com/en-us/graph/api/serviceprincipal-get?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/serviceprincipal-get?view=graph-rest-1.0)
