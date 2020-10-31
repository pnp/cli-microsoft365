# spo field set

Updates existing list or site column

## Usage

```sh
m365 spo field set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the field is located

`--listId [listId]`
: ID of the list where the field is located (if list column). Specify `listTitle` or `listId` but not both

`--listTitle [listTitle]`
: Title of the list where the field is located (if list column). Specify `listTitle` or `listId` but not both

`-i
: --id [id]`|ID of the field to update. Specify `name` or `id` but not both

`-n
: --name [name]`|Title or internal name of the field to update. Specify `name` or `id` but not both

`--updateExistingLists`
: Set, to push the update to existing lists. Otherwise, the changes will apply to new lists only

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Specify properties to update using their names, eg. `--Title 'New Title' --JSLink jslink.js`.

## Examples

Update the title of the site column specified by its internal name and push changes to existing lists

```sh
m365 spo field set --webUrl https://contoso.sharepoint.com/sites/project-x --name 'MyColumn' --updateExistingLists --Title 'My column'
```

Update the title of the list column specified by its ID

```sh
m365 spo field set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --id 330f29c5-5c4c-465f-9f4b-7903020ae1ce --Title 'My column'
```

Update column formatting of the specified list column

```sh
m365 spo field set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --name 'MyColumn' --CustomFormatter '`{"schema":"https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json", "elmType": "div", "txtContent": "@currentField"}`'
```