# login

Log in to Microsoft 365

## Usage

```sh
login [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-t, --authType [authType]`|The type of authentication to use. Allowed values `certificate,deviceCode,password,identity`. Default `deviceCode`
`-u, --userName [userName]`|Name of the user to authenticate. Required when `authType` is set to `password`
`-p, --password [password]`|Password for the user. Required when `authType` is set to `password`
`-c, --certificateFile [certificateFile]`|Path to the file with certificate private key. Required when `authType` is set to `certificate`
`--thumbprint [thumbprint]`|Certificate thumbprint. Required when `authType` is set to `certificate`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Using the `login` command you can log in to Microsoft 365.

By default, the `login` command uses device code OAuth flow to log in to Microsoft 365. Alternatively, you can authenticate using a user name and password or certificate, which are convenient for CI/CD scenarios, but which come with their own limitations. See the CLI for Microsoft 365 manual for more information.

When logging in to Microsoft 365, the `login` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [logout](logout.md) command.

When logging in to Microsoft 365 using the user name and password, next to the access and refresh token, the CLI for Microsoft 365 will store the user credentials so that it can automatically re-authenticate if necessary. Similarly to the tokens, the credentials are removed by re-authenticating using the device code or by calling the [logout](logout.md) command.

When logging in to Microsoft 365 using a certificate, the CLI for Microsoft 365 will store the contents of the certificate so that it can automatically re-authenticate if necessary. The contents of the certificate are removed by re-authenticating using the device code or by calling the [logout](logout.md) command.

To log in to Microsoft 365 using a certificate, you will typically create a custom Azure AD application. To use this application with the CLI for Microsoft 365, you will set the `CLIMICROSOFT365_AADAPPID` environment variable to the application's ID and the `CLIMICROSOFT365_TENANT` environment variable to the ID of the Azure AD tenant, where you created the Azure AD application.

Managed identity in Azure Cloud Shell is the identity of the user. It is neither system- nor user-assigned and it can't be configured. To log in to Microsoft 365 using managed identity in Azure Cloud Shell, set `authType` to `identity` and don't specify the `userName` option.

## Examples

Log in to Microsoft 365 using the device code

```sh
login
```

Log in to Microsoft 365 using the device code in debug mode including detailed debug information in the console output

```sh
login --debug
```

Log in to Microsoft 365 using a user name and password

```sh
login --authType password --userName user@contoso.com --password pass@word1
```

Log in to Microsoft 365 using a PEM certificate

```sh
login --authType certificate --certificateFile /Users/user/dev/localhost.pem --thumbprint 47C4885736C624E90491F32B98855AA8A7562AF1
```

Log in to Microsoft 365 using a personal information exchange (.pfx) file

```sh
login --authType certificate --certificateFile /Users/user/dev/localhost.pfx --thumbprint 47C4885736C624E90491F32B98855AA8A7562AF1 --password 'pass@word1'
```

Log in to Microsoft 365 using a system assigned managed identity. Applies to Azure resources with managed identity enabled,
such as Azure Virtual Machines, Azure App Service or Azure Functions

```sh
login --authType identity
```

Log in to Microsoft 365 using managed identity in Azure Cloud Shell. Uses the identity of the current user.

```sh
login --authType identity
```

Log in to Microsoft 365 using a user-assigned managed identity. Client id or principal id also known as object id value can be specified in the `userName` option. Applies to Azure resources with managed identity enabled, such as Azure Virtual Machines, Azure App Service or Azure Functions

```sh
login --authType identity --userName ac9fbed5-804c-4362-a369-21a4ec51109e
```
