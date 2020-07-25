# spo site commsite enable

Enables communication site features on the specified site

## Usage

```sh
m365 spo site commsite enable [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The URL of the site to enable communication site features on

`-i, --designPackageId [designPackageId]`
: The ID of the site design to apply when enabling communication site features

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

Enable communication site features on an existing site

```sh
m365 spo site commsite enable --url https://contoso.sharepoint.com
```