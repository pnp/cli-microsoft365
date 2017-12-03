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
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Arguments

Argument|Description
--------|-----------
`url`|absolute URL of the SharePoint Online site to connect to

## Remarks

Using the `spo connect` command, you can connect to any SharePoint Online site.
Depending on the command you want to use, you might be required to connect
to a SharePoint Online tenant admin site (suffixed with _-admin_,
eg. _https://contoso-admin.sharepoint.com_) or a regular site.

The `spo connect` command uses device code OAuth flow with the standard
Microsoft SharePoint Online Management Shell Azure AD application to connect
to SharePoint Online.

When connecting to a SharePoint site, the `spo connect` command stores in memory
the access token and the refresh token for the specified site. Both tokens are cleared from memory
after exiting the CLI or by calling the [spo disconnect](connect.md) command.

## Examples

Connects to a SharePoint Online tenant admin site

```sh
spo connect https://contoso-admin.sharepoint.com
```

Connect to a SharePoint Online tenant admin site in debug mode including detailed debug information in the console output

```sh
spo connect --debug https://contoso-admin.sharepoint.com
```

Connect to a regular SharePoint Online site

```sh
spo connect https://contoso.sharepoint.com/sites/team
```