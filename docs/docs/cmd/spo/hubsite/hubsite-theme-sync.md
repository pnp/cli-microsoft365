# spo hubsite theme sync

Applies any theme updates from the parent hub site.

## Usage

```sh
m365 spo hubsite theme sync [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site to apply theme updates from the hub site to.

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

## Examples

Applies any theme updates from the parent hub site to the site with URL https://contoso.sharepoint.com/sites/project-x

```sh
m365 spo hubsite theme sync --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)