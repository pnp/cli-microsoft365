# spo tenant recyclebinitem restore

Restores the specified deleted Site Collection from Tenant Recycle Bin

## Usage

```sh
spo tenant recyclebinitem restore [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url`|URL of the site to restore
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Restore a deleted site collection from tenant recycle bin

```sh
spo tenant recyclebinitem restore --url https://contoso.sharepoint.com/sites/team
```
