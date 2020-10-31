# spo hubsite list

Lists hub sites in the current tenant

## Usage

```sh
m365 spo hubsite list [options]
```

## Options

`-h, --help`
: output usage information

`-i, --includeAssociatedSites`
: Include the associated sites in the result (only in JSON output)

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

When using the text output type (default), the command lists only the values of the `ID`, `SiteUrl` and `Title` properties of the hub site. When setting the output type to JSON, all available properties are included in the command output.

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