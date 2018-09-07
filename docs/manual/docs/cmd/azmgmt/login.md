# azmgmt login

Log in to the Azure Management Service

## Usage

```sh
azmgmt login [options]
```

## Alias

```sh
azmgmt connect
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
    The 'azmgmt connect' command is deprecated. Please use 'azmgmt login' instead.

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

Using the `azmgmt login` command you can log in to the Azure Management Service to manage your Azure objects.

By default, the `azmgmt login` command uses device code OAuth flow to log in to the Azure Management Service. Alternatively, you can authenticate using a user name and password, which is convenient for CI/CD scenarios, but which comes with its own limitations. See the Office 365 CLI manual for more information.

When logging in to the Azure Management Service, the `azmgmt login` command stores in memory the access token and the refresh token. Both tokens are cleared from memory after exiting the CLI or by calling the [azmgmt logout](logout.md) command.

When logging in to the Azure Management Service using the user name and password, next to the access and refresh token, the Office 365 CLI will store the user credentials so that it can automatically reauthenticate if necessary. Similarly to the tokens, the credentials are removed by reauthenticating using the device code or by calling the azmgmt logout command.

## Examples

Log in to the Azure Management Service using the device code

```sh
azmgmt login
```

Log in to the Azure Management Service using the device code in debug mode including detailed debug information in the console output

```sh
azmgmt login --debug
```

Log in to the Azure Management Service using a user name and password

```sh
azmgmt login --authType password --userName user@contoso.com --password pass@word1
```