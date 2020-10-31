# spo list webhook remove

Removes the specified webhook from the list

## Usage

```sh
m365 spo list webhook remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to remove the webhook from is located

`-l, --listId [listId]`
: ID of the list from which the webhook should be removed. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list from which the webhook should be removed. Specify either `listId` or `listTitle` but not both

`-i, --id <id>`
: ID of the webhook to remove

`--confirm`
: Don't prompt for confirming removing the webhook

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified id doesn't refer to an existing webhook, you will get a `404 - "404 FILE NOT FOUND"` error.

## Examples

Remove webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ from a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/ninja_

```sh
m365 spo list webhook remove --webUrl https://contoso.sharepoint.com/sites/ninja --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id cc27a922-8224-4296-90a5-ebbc54da2e81
```

Remove webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ from a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/ninja_

```sh
m365 spo list webhook remove --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e81
```

Remove webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ from a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/ninja_ without being asked for confirmation

```sh
m365 spo list webhook remove --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e81 --confirm
```