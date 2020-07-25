# spo orgnewssite remove

Removes a site from the list of organizational news sites

## Usage

```sh
m365 spo orgnewssite remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: Absolute URL of the site to remove

`--confirm`
: Don't prompt for confirmation

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

## Examples

Remove a site from the list of organizational news

```sh
m365 spo orgnewssite remove --url https://contoso.sharepoint.com/sites/site1
```

Remove a site from the list of organizational news sites, without prompting for confirmation

```sh
m365 spo orgnewssite remove --url https://contoso.sharepoint.com/sites/site1 --confirm
```