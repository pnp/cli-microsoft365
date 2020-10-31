# spo list label set

Sets classification label on the specified list

## Usage

```sh
m365 spo list label set  [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the list is located

`--label <label>`
: The label to set on the list

`-t, --listTitle [listTitle]`
: The title of the list on which to set the label. Specify only one of `listTitle`, `listId` or `listUrl`

`-l, --listId [listId]`
: The ID of the list on which to set the label. Specify only one of `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or web-relative URL of the list on which to set the label. Specify only one of `listTitle`, `listId` or `listUrl`

`--syncToItems`
: Specify, to set the label on all items in the list

`--blockDelete`
: Specify, to disallow deleting items in the list

`--blockEdit`
: Specify, to disallow editing items in the list

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Sets classification label "Confidential" for list _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list label set --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Shared Documents' --label 'Confidential'
```

Sets classification label "Confidential" and disables editing and deleting items on the list and all existing items for list for list _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list label set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'Documents' --label 'Confidential' --blockEdit --blockDelete --syncToItems
```