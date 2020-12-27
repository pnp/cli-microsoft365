# spo site get

Gets information about the specific site collection

## Usage

```sh
m365 spo site get [options]
```

## Options

`-u, --url <url>`
: URL of the site collection to retrieve information for

--8<-- "docs/cmd/_global.md"

## Remarks

This command can retrieve information for both classic and modern sites.

## Examples

Return information about the _https://contoso.sharepoint.com/sites/project-x_ site collection.

```sh
m365 spo site get -u https://contoso.sharepoint.com/sites/project-x
```
