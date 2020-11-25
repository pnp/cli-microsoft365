# spo tenant recyclebinitem restore

Restores the specified deleted Site Collection from Tenant Recycle Bin

## Usage

```sh
m365 spo tenant recyclebinitem restore [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url`
: URL of the site to restore

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

Restoring deleted site collections is by default asynchronous and depending on the current state of Microsoft 365, might take up to few minutes. If you're building a script with steps that require the site to be fully restored, you should use the `--wait` flag. When using this flag, the `spo tenant recyclebinitem restore` command will keep running until it received confirmation from Microsoft 365 that the site has been fully restored.

## Examples

Restore a deleted site collection from tenant recycle bin

```sh
m365 spo tenant recyclebinitem restore --url https://contoso.sharepoint.com/sites/team
```

Restore a deleted site collection from tenant recycle bin and wait for completion

```sh
m365 spo tenant recyclebinitem restore --url https://contoso.sharepoint.com/sites/team --wait
```
