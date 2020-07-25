# spo list contenttype list

Lists content types configured on the list

## Usage

```sh
m365 spo list contenttype list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-l, --listId [listId]`
: ID of the list for which to list configured content types. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list for which to list configured content types. Specify `listId` or `listTitle` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

List all content types configured on the list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype list --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```

List all content types configured on the list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list contenttype list --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents
```