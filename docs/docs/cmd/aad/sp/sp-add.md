# aad sp add

Adds a service principal to a registered Azure AD app

## Usage

```sh
m365 aad sp add [options]
```

## Options

`--appId [appId]`
: ID of the application to which the service principal should be added

`--appName [appName]`
: Display name of the application to which the service principal should be added

`--objectId [objectId]`
: ObjectId of the application to which the service principal should be added

--8<-- "docs/cmd/_global.md"

## Remarks

Specify either the `appId`, `appName` or `objectId`. If you specify more than one option value, the command will fail with an error.

If you register an application in the portal, an application object as well as a service principal object are automatically created in your home tenant. If you register an application using CLI for Microsoft 365 or the Microsoft Graph, you'll need to create the service principal separately. To register/create an application using the CLI for Microsoft 365, use the [m365 aad app add](../app/app-add.md) command.

## Examples

Adds a service principal to a registered Azure AD app with appId _b2307a39-e878-458b-bc90-03bc578531d6_.

```sh
m365 aad sp add --appId b2307a39-e878-458b-bc90-03bc578531d6
```

Adds a service principal to a registered Azure AD app with appName _Microsoft Graph_.

```sh
m365 aad sp add --appName "Microsoft Graph"
```

Adds a service principal to a registered Azure AD app with objectId _b2307a39-e878-458b-bc90-03bc578531d6_.

```sh
m365 aad sp add --objectId b2307a39-e878-458b-bc90-03bc578531d6
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)
- Create servicePrincipal: [https://docs.microsoft.com/en-us/graph/api/serviceprincipal-post-serviceprincipals?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/serviceprincipal-post-serviceprincipals?view=graph-rest-1.0)
