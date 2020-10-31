# spo list view list

Lists views configured on the specified list

## Usage

```sh
m365 spo list view list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-i, --listId [listId]`
: ID of the list for which to list configured views. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list for which to list configured views. Specify `listId` or `listTitle` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

List all views for a list with title *Documents* located in site *https://contoso.sharepoint.com/sites/project-x*

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents
```

List all views for a list with ID *0cd891ef-afce-4e55-b836-fce03286cccf* located in site *https://contoso.sharepoint.com/sites/project-x*

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```
