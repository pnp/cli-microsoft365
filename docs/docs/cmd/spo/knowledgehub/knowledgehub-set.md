# spo knowledgehub set

Sets the Knowledge Hub Site for your tenant

## Usage

```sh
m365 spo knowledgehub set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: URL of the site to set as Knowledge Hub

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

If the specified url doesn't refer to an existing site collection, you will get a `404 - "404 FILE NOT FOUND"` error.

## Examples

Sets the Knowledge Hub Site for your tenant

```sh
m365 spo knowledgehub set --url https://contoso.sharepoint.com/sites/knowledgesite
```
