# login

Log in to Microsoft 365

## Usage

```sh
m365 login [options]
```

## Options

`-t, --authType [authType]`
: The type of authentication to use. Allowed values `certificate,deviceCode,password,identity,browser`. Default `deviceCode`

`-u, --userName [userName]`
: Name of the user to authenticate. Required when `authType` is set to `password`

`-p, --password [password]`
: Password for the user or the certificate. Required when `authType` is set to `password`, or when `authType` is set to `certificate` and the provided certificate requires a password to open

`-c, --certificateFile [certificateFile]`
: Path to the file with certificate private key. When `authType` is set to `certificate`, specify either `certificateFile` or `certificateBase64Encoded`

`--certificateBase64Encoded [certificateBase64Encoded]`
: Base64-encoded string with certificate private key. When `authType` is set to `certificate`, specify either `certificateFile` or `certificateBase64Encoded`

`--thumbprint [thumbprint]`
: Certificate thumbprint. If not specified, and `authType` is set to `certificate`, it will be automatically calculated based on the specified certificate

`--appId [appId]`
: App ID of the Azure AD application to use for authentication. If not specified, use the app specified in the `CLIMICROSOFT365_AADAPPID` environment variable. If the environment variable is not defined, use the multitenant PnP Management Shell app

`--tenant [tenant]`
: ID of the tenant from which accounts should be able to authenticate. Use `common` or `organization` if the app is multitenant. If not specified, use the tenant specified in the `CLIMICROSOFT365_TENANT` environment variable. If the environment variable is not defined, use `common` as the tenant identifier

--8<-- "docs/cmd/_global.md"

## Remarks

Using the `login` command you can log in to Microsoft 365.

By default, the `login` command uses device code OAuth flow to log in to Microsoft 365. Alternatively, you can authenticate using a user name and password or certificate, which are convenient for CI/CD scenarios, but which come with their own [limitations](../user-guide/connecting-office-365.md).

When logging in to Microsoft 365, the `login` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [logout](logout.md) command.

When logging in to Microsoft 365 using the user name and password, next to the access and refresh token, the CLI for Microsoft 365 will store the user credentials so that it can automatically re-authenticate if necessary. Similarly to the tokens, the credentials are removed by re-authenticating using the device code or by calling the [logout](logout.md) command.

When logging in to Microsoft 365 using a certificate, the CLI for Microsoft 365 will store the contents of the certificate so that it can automatically re-authenticate if necessary. The contents of the certificate are removed by re-authenticating using the device code or by calling the [logout](logout.md) command.  

To log in to Microsoft 365 using a certificate, you will typically [create a custom Azure AD application](../user-guide/using-own-identity.md). To use this application with the CLI for Microsoft 365, you will set the `CLIMICROSOFT365_AADAPPID` environment variable to the application's ID and the `CLIMICROSOFT365_TENANT` environment variable to the ID of the Azure AD tenant, where you created the Azure AD application. Also, please make sure to read about [the caveats when using the certificate login option](../user-guide/cli-certificate-caveats.md).

Managed identity in Azure Cloud Shell is the identity of the user. It is neither system- nor user-assigned and it can't be configured. To log in to Microsoft 365 using managed identity in Azure Cloud Shell, set `authType` to `identity` and don't specify the `userName` option.

## Examples

Log in to Microsoft 365 using the device code

```sh
m365 login
```

Log in to Microsoft 365 using the device code in debug mode including detailed debug information in the console output

```sh
m365 login --debug
```

Log in to Microsoft 365 using a user name and password

```sh
m365 login --authType password --userName user@contoso.com --password pass@word1
```

Log in to Microsoft 365 using a PEM certificate

```sh
m365 login --authType certificate --certificateFile /Users/user/dev/localhost.pem
```

Log in to Microsoft 365 using a PEM certificate. Use the specified thumbprint

```sh
m365 login --authType certificate --certificateFile /Users/user/dev/localhost.pem  --thumbprint 47C4885736C624E90491F32B98855AA8A7562AF1
```

Log in to Microsoft 365 using a personal information exchange (.pfx) file

```sh
m365 login --authType certificate --certificateFile /Users/user/dev/localhost.pfx --password 'pass@word1'
```

Log in to Microsoft 365 using a personal information exchange (.pfx) file protected with an empty password

```sh
m365 login --authType certificate --certificateFile /Users/user/dev/localhost.pfx --password
```

Log in to Microsoft 365 using a personal information exchange (.pfx) file not protected with a password

```sh
m365 login --authType certificate --certificateFile /Users/user/dev/localhost.pfx
```

Log in to Microsoft 365 using a personal information exchange (.pfx) file. Use the specified thumbprint

```sh
m365 login --authType certificate --certificateFile /Users/user/dev/localhost.pfx --thumbprint 47C4885736C624E90491F32B98855AA8A7562AF1 --password 'pass@word1'
```

Log in to Microsoft 365 using a certificate from a base64-encoded string

```sh
m365 login --authType certificate --certificateBase64Encoded MIII2QIBAzCCCJ8GCSqGSIb3DQEHAaCCCJAEg...eX1N5AgIIAA== --thumbprint D0C9B442DE249F55D10CDA1A2418952DC7D407A3
```

Log in to Microsoft 365 using a system assigned managed identity. Applies to Azure resources with managed identity enabled,
such as Azure Virtual Machines, Azure App Service or Azure Functions

```sh
m365 login --authType identity
```

Log in to Microsoft 365 using managed identity in Azure Cloud Shell. Uses the identity of the current user.

```sh
m365 login --authType identity
```

Log in to Microsoft 365 using a user-assigned managed identity. Client id or principal id also known as object id value can be specified in the `userName` option. Applies to Azure resources with managed identity enabled, such as Azure Virtual Machines, Azure App Service or Azure Functions

```sh
m365 login --authType identity --userName ac9fbed5-804c-4362-a369-21a4ec51109e
```

Log in to Microsoft 365 using your own multitenant Azure AD application

```sh
m365 login --appId 31359c7f-bd7e-475c-86db-fdb8c937548c
```

Log in to Microsoft 365 using your own Azure AD application that's restricted only to allow accounts from the specific tenant

```sh
m365 login --appId 31359c7f-bd7e-475c-86db-fdb8c937548c --tenant 31359c7f-bd7e-475c-86db-fdb8c937548a
```

Log in to Microsoft 365 using the interactive browser authentication. Uses the identity of the current user.

```sh
m365 login --authType browser
```
