# spo hubsite set

Updates properties of the specified hub site

## Usage

```sh
m365 spo hubsite set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: ID of the hub site to update

`-t, --title [title]`
: The new title for the hub site

`-d, --description [description]`
: The new description for the hub site

`-l, --logoUrl [logoUrl]`
: The URL of the new logo for the hub site

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

If the specified `id` doesn't refer to an existing hub site, you will get an `Unknown Error` error.

## Examples

Update hub site's title

```sh
m365 spo hubsite set --id 255a50b2-527f-4413-8485-57f4c17a24d1 --title Sales
```

Update hub site's title and description

```sh
m365 spo hubsite set --id 255a50b2-527f-4413-8485-57f4c17a24d1 --title Sales --description "All things sales"
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)