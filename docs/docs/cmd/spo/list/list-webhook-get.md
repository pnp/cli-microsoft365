# spo list webhook get

Gets information about the specific webhook

## Usage

```sh
m365 spo list webhook get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to retrieve the webhook info for is located

`-l, --listId [listId]`
: ID of the list from which to retrieve the webhook. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list from which to retrieve the webhook. Specify either `listId` or `listTitle` but not both

`-i, --id [id]`
: ID of the webhook to retrieve

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified `id` doesn't refer to an existing webhook, you will get a `404 - "404 FILE NOT FOUND"` error.

## Examples

Return information about a webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e85_ which belongs to a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list webhook get --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id cc27a922-8224-4296-90a5-ebbc54da2e85
```

Return information about a webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e85_ which belongs to a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list webhook get --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e85
```