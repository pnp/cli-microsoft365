# spo list webhook list

Lists all webhooks for the specified list

## Usage

```sh
m365 spo list webhook list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to retrieve webhooks for is located

`-i, --listId [listId]`
: ID of the list to retrieve all webhooks for. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list to retrieve all webhooks for. Specify either `listId` or `listTitle` but not both

`--id [id]`
: (deprecated. Use `listId` instead) ID of the list to retrieve all webhooks for. Specify either `id` or `title` but not both

`--title [title]`
: (deprecated. Use `listTitle` instead) Title of the list to retrieve all webhooks for. Specify either `id` or `title` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

List all webhooks for a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list webhook list --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```

List all webhooks for a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list webhook list --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents
```