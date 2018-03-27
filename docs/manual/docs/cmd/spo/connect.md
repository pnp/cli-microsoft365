# spo connect

Connects to a SharePoint Online site

## Usage

```sh
spo connect [options] <url>
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

## Arguments

Argument|Description
--------|-----------
`url`|absolute URL of the SharePoint Online site to connect to

## Remarks

Using the `spo connect` command, you can connect to any SharePoint Online site. Depending on the command you want to use, you might be required to connect to a SharePoint Online tenant admin site (suffixed with _-admin_, eg. _https://contoso-admin.sharepoint.com_) or a regular site.

By default, the `spo connect` command uses device code OAuth flow to connect to SharePoint Online. Alternatively, you can authenticate using a user name and password, which is convenient for CI/CD scenarios, but which comes with its own limitations. See the Office 365 CLI manual for more information.

When connecting to a SharePoint site, the `spo connect` command stores in memory the access token and the refresh token for the specified site. Both tokens are cleared from memory after exiting the CLI or by calling the [spo disconnect](connect.md) command.

When connecting to SharePoint Online using the user name and password, next to the access and refresh token, the Office 365 CLI will store the user credentials so that it can automatically reauthenticate if necessary. Similarly to the tokens, the credentials are removed by reconnecting using the device code or by calling the `spo disconnect` command.

## Examples

Connects to a SharePoint Online tenant admin site using the device code

```sh
spo connect https://contoso-admin.sharepoint.com
```

Connect to a SharePoint Online tenant admin site using the device code in debug mode including detailed debug information in the console output

```sh
spo connect --debug https://contoso-admin.sharepoint.com
```

Connect to a regular SharePoint Online site using the device code

```sh
spo connect https://contoso.sharepoint.com/sites/team
```

Connect to a SharePoint Online tenant admin site using a user name and password

```sh
spo connect https://contoso-admin.sharepoint.com --authType password --userName user@contoso.com --password pass@word1
```