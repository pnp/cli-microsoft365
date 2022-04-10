# spo get

Gets the context URL for the root SharePoint site collection and SharePoint tenant admin site

## Usage

```sh
m365 spo get [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

CLI for Microsoft 365 automatically discovers the URL of the root SharePoint site collection/SharePoint tenant admin site (whichever is needed to run the particular command). Using this command you can see which URLs the CLI has discovered.

## Examples

Get the context URL for the root SharePoint site collection and SharePoint tenant admin site

```sh
m365 spo get --output json
```
