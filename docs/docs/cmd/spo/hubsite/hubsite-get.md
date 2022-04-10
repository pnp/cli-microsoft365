# spo hubsite get

Gets information about the specified hub site

## Usage

```sh
m365 spo hubsite get [options]
```

## Options

`-i, --id <id>`
: Hub site ID

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

If the specified `id` doesn't refer to an existing hub site, you will get a `ResourceNotFoundException` error.

## Examples

Get information about the hub site with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
m365 spo hubsite get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
