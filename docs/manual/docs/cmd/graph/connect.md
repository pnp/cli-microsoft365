# graph connect

Connects to the Microsoft Graph

## Usage

```sh
graph connect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-t, --authType [authType]`|The type of authentication to use. Allowed values `deviceCode|password|certificate`. Default `deviceCode`
`-u, --userName [userName]`|Name of the user to authenticate. Required when `authType` is set to `password`
`-p, --password [password]`|Password for the user. Required when `authType` is set to `password`
`-c, --certificateFile [certificateFile]`|File name for file with private certificate key. Required when `authType` is set to `certificate`
`-a, --applicationId [applicationId]`|Application ID of Azure AD app. Required when `authType` is set to `certificate`
`--thumbprint [thumbprint]`|Certificate thumbprint. Required when `authType` is set to `certificate`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Using the `graph connect` command you can connect to the Microsoft Graph.

By default, the `graph connect` command uses device code OAuth flow to connect to the Microsoft Graph. Alternatively, you can authenticate using a user name and password or a certificate, which is convenient for CI/CD scenarios, but which comes with its own limitations. See the Office 365 CLI manual for more information.

When connecting to the Microsoft Graph, the `graph connect` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [graph disconnect](disconnect.md) command.

When connecting to the Microsoft Graph using the user name and password, next to the access and refresh token, the Office 365 CLI will store the user credentials so that it can automatically reauthenticate if necessary. Similarly to the tokens, the credentials are removed by reconnecting using the device code or by calling the `graph disconnect` command.

When connecting to the Microsoft Graph using a certificate you are required to create your own Azure AD application and grant permissions accordingly. You are alsoo required to specify the `OFFICE365CLI_TENANT` environment variable which should have the value of your tenant name; for instance `contoso.onmicrosoft.com`. Not all commands will work with a certificate as not all features in the Microsoft Graph supports App-only policies.

## Examples

Connect to the Microsoft Graph using the device code

```sh
graph connect
```

Connect to the Microsoft Graph using the device code in debug mode including detailed debug information in the console output

```sh
graph connect --debug
```

Connect to the Microsoft Graph using a user name and password

```sh
graph connect --authType password --userName user@contoso.com --password pass@word1
```

Connect to the Microsoft Graph using a certificate

```sh
graph connect --authType certificate --certificateFile cert.pem --thumbprint d712ebab09e3a9788e9d1a234ea4ac98d173c6c3 --clientId b269214b-7ed2-4d60-9fb2-064c7b79a4a3
```