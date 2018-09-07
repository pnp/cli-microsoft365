# aad login

Log in to the Azure Active Directory Graph

## Usage

```sh
aad login [options]
```

## Alias

```sh
aad connect
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-t, --authType [authType]`|The type of authentication to use. Allowed values `deviceCode|password`. Default `deviceCode`
`-u, --userName [userName]`|Name of the user to authenticate. Required when `authType` is set to `password`
`-p, --password [password]`|Password for the user. Required when `authType` is set to `password`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    The 'aad connect' command is deprecated. Please use 'aad login' instead.

Using the `aad login` command you can log in to the Azure Active Directory Graph to manage your AAD objects.

By default, the `aad login` command uses device code OAuth flow to log in to the AAD Graph. Alternatively, you can authenticate using a user name and password, which is convenient for CI/CD scenarios, but which comes with its own limitations. See the Office 365 CLI manual for more information.

When logging in to the AAD Graph, the `aad login` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [aad logout](logout.md) command.

When logging in to the AAD Graph using the user name and password, next to the access and refresh token, the Office 365 CLI will store the user credentials so that it can automatically reauthenticate if necessary. Similarly to the tokens, the credentials are removed by reauthenticating using the device code or by calling the `aad logout` command.

## Examples

Log in to the AAD Graph using the device code

```sh
aad login
```

Log in to the AAD Graph using the device code in debug mode including detailed debug information in the console output

```sh
aad login --debug
```

Log in to the AAD Graph using a user name and password

```sh
aad login --authType password --userName user@contoso.com --password pass@word1
```