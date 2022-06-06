# spo list view add

Adds a new view to a SharePoint list

## Usage

```sh
m365 spo list view add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located.

`--listId [listId]`
: ID of the list to which the view should be added. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--listTitle [listTitle]`
: Title of the list to which the view should be added. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--listUrl [listUrl]`
: Relative URL of the list to which the view should be added. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--title <title>`
: Title of the view to be created for the list.

`--fields <fields>`
: Comma-separated list of **case-sensitive** internal names of the fields to add to the view.

`--viewQuery [viewQuery]`
: XML representation of the list query for the underlying view.

`--personal`
: View will be created as personal view, if specified.

`--default`
: View will be set as default view, if specified.

`--paged`
: View supports paging, if specified (recommended to use this).

`--rowLimit [rowLimit]`
: Sets the number of items to display for the view. Default value is 30.

--8<-- "docs/cmd/_global.md"

## Remarks

We recommend using the `paged` option. When specified, the view supports displaying more items page by page (default behavior). When not specified, the `rowLimit` is absolute, and there is no link to see more items.

## Examples

Add a view called _All events_ to a list with specific title.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "My List" --title "All events" --fields "FieldName1,FieldName2,Created,Author,Modified,Editor" --paged
```

Add a view as default view with title _All events_ to a list with a specific URL.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl "/Lists/MyList" --title "All events" --fields "FieldName1,Created" --paged --default
```

Add a personal view called _All events_ to a list with a specific ID.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listId 00000000-0000-0000-0000-000000000000 --title "All events" --fields "FieldName1,Created" --paged --personal
```

Add a view called _All events_ with defined filter and sorting.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "My List" --title "All events" --fields "FieldName1" --viewQuery "<OrderBy><FieldRef Name='Created' Ascending='FALSE' /></OrderBy><Where><Eq><FieldRef Name='TextFieldName' /><Value Type='Text'>Field value</Value></Eq></Where>" --paged
```
