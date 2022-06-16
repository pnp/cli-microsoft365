# spo hubsite list

Lists hub sites in the current tenant

## Usage

```sh
m365 spo hubsite list [options]
```

## Options

`-i, --includeAssociatedSites`
: Include the associated sites in the result (only in JSON output)

--8<-- "docs/cmd/_global.md"

## Remarks

When using the text or csv output type, the command lists only the values of the `ID`, `SiteUrl` and `Title` properties of the hub site. With the output type as JSON, all available properties are included in the command output.

## Examples

List hub sites in the current tenant

```sh
m365 spo hubsite list
```

List hub sites, including their associated sites, in the current tenant. Associated site info is only shown in JSON output.

```sh
m365 spo hubsite list --includeAssociatedSites --output json
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
