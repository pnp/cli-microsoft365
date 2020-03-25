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
`-u, --url`|URL of the Site Collection to restore
`--wait`|Wait for the Site Collection to be restored before completing the command
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Restore a deleted Site Collection from Tenant Recycle Bin

```sh
spo tenant recyclebinitem restore --url https://contoso.sharepoint.com/sites/team
```

Restore a deleted Site Collection from Tenant Recycle Bin and wait for the restoring process to complete

```sh
spo tenant recyclebinitem restore --url https://contoso.sharepoint.com/sites/team --wait
```
