# spo orgnewssite set

Marks site as an organizational news site

## Usage

```sh
spo orgnewssite set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|The URL of the site to mark as an organizational news site
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-u, --url` option you can specify which site to add to the list of organizational news sites.

## Examples

Set a site as an organizational news site

```sh
spo orgnewssite set --url https://contoso.sharepoint.com/sites/site1
```
