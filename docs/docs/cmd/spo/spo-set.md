# spo set

Sets the URL of the root SharePoint site collection for use in SPO commands

## Usage

```sh
m365 spo set [options]
```

## Options

`-u, --url <url>`
: The URL of the root SharePoint site collection to use in SPO commands

--8<-- "docs/cmd/_global.md"

## Remarks

CLI for Microsoft 365 automatically discovers the URL of the root SharePoint site collection/SharePoint tenant admin site (whichever is needed to run the particular command). In specific cases, like when managing multi-geo Microsoft 365 tenants, it could be desirable to make the CLI manage the specific geography. For such cases, you can use this command to explicitly specify the SPO URL that should be used when executing SPO commands.

## Examples

Set SPO URL to the specified URL

```sh
m365 spo set --url https://contoso.sharepoint.com
```
