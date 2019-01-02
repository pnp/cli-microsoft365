# graph login

Log in to the Microsoft Graph

## Usage

```sh
graph login [options]
```

## Alias

```sh
graph connect
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-t, --authType [authType]`|The type of authentication to use. Allowed values `certificate|deviceCode|password`. Default `deviceCode`
`-u, --userName [userName]`|Name of the user to authenticate. Required when `authType` is set to `password`
`-p, --password [password]`|Password for the user. Required when `authType` is set to `password`
`-c, --certificateFile [certificateFile]`|Path to the file with certificate private key. Required when `authType` is set to `certificate`
`--thumbprint [thumbprint]`|Certificate thumbprint. Required when `authType` is set to `certificate`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    The 'graph connect' command is deprecated. Please use 'graph login' instead.

Using the `graph login` command you can log in to the Microsoft Graph.

By default, the `graph login` command uses device code OAuth flow to log in to the Microsoft Graph. Alternatively, you can authenticate using a user name and password or certificate, which are convenient for CI/CD scenarios, but which come with their own limitations. See the Office 365 CLI manual for more information.

When logging in to the Microsoft Graph, the `graph login` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [graph logout](logout.md) command.

When logging in to the Microsoft Graph using the user name and password, next to the access and refresh token, the Office 365 CLI will store the user credentials so that it can automatically re-authenticate if necessary. Similarly to the tokens, the credentials are removed by re-authenticating using the device code or by calling the [graph logout](logout.md) command.

When logging in to the Microsoft Graph using a certificate, the Office 365 CLI will store the contents of the certificate so that it can automatically re-authenticate if necessary. The contents of the certificate are removed by re-authenticating using the device code or by calling the [graph logout](logout.md) command.

To log in to the Microsoft Graph using a certificate, you will typically create a custom Azure AD application. To use this application with the Office 365 CLI, you will set the `OFFICE365CLI_AADAADAPPID` environment variable to the application's ID and the `OFFICE365CLI_TENANT` environment variable to the ID of the Azure AD tenant, where you created the Azure AD application.

## Examples

Log in to the Microsoft Graph using the device code

```sh
graph login
```

Log in to the Microsoft Graph using the device code in debug mode including detailed debug information in the console output

```sh
graph login --debug
```

Log in to the Microsoft Graph using a user name and password

```sh
graph login --authType password --userName user@contoso.com --password pass@word1
```

Log in to the Microsoft Graph using a certificate

```sh
graph login --authType certificate --certificateFile /Users/user/dev/localhost.pfx --thumbprint 47C4885736C624E90491F32B98855AA8A7562AF1
```