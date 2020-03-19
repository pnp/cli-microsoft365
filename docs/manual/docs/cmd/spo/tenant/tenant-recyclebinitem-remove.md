# spo tenant recyclebinitem remove

Removes the specified deleted Site Collection from Tenant Recycle Bin

## Usage

```sh
spo tenant recyclebinitem remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url`|URL of the Site Collection to remove
`--wait`|Wait for the Site Collection to be removed before completing the command
`--confirm`|Don't prompt for confirming removing the deleted Site Collection
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Removes the specified deleted Site Collection from Tenant Recycle Bin

```sh
spo tenant recyclebinitem remove --url https://contoso.sharepoint.com/sites/team
```

Removes the specified deleted Site Collection from Tenant Recycle Bin and wait for the removing process to complete

```sh
spo tenant recyclebinitem remove --url https://contoso.sharepoint.com/sites/team --wait
```