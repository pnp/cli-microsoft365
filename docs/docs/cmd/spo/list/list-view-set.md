# spo list view set

Updates existing list view

## Usage

```sh
m365 spo list view set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`--listId [listId]`
: ID of the list where the view is located. Specify `listTitle` or `listId` but not both

`--listTitle [listTitle]`
: Title of the list where the view is located. Specify `listTitle` or `listId` but not both

`--viewId [viewId]`
: ID of the view to update. Specify `viewTitle` or `viewId` but not both

`--viewTitle [viewTitle]`
: Title of the view to update. Specify `viewTitle` or `viewId` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Specify properties to update using their names, eg. `--Title 'New Title' --JSLink jslink.js`

When updating list formatting, the value of the CustomFormatter property must be XML-escaped, eg. `&lt;` instead of `<`.

## Examples

Update the title of the list view specified by its name

```sh
m365 spo list view set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewTitle 'All items' --Title 'All events'
```

Update the title of the list view specified by its ID

```sh
m365 spo list view set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewId 330f29c5-5c4c-465f-9f4b-7903020ae1ce --Title 'All events'
```

Update view formatting of the specified list view

```sh
m365 spo list view set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewTitle 'All items' --CustomFormatter '`{"schema":"https://developer.microsoft.com/json-schemas/sp/view-formatting.schema.json","additionalRowClass": "=if([$DueDate] &lt;= @now, 'sp-field-severity--severeWarning', '')"}`'
```