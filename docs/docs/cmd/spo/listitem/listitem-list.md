# spo listitem list

Gets a list of items from the specified list

## Usage

```sh
m365 spo listitem list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved

`-i, --listId [listId]`
: ID of the list to retrieve items from. Specify `listId` or `listTitle` but not both

`--id [id]`
: (deprecated. Use `listId` instead) ID of the list to retrieve items from. Specify `id` or `title` but not both

`-t, --listTitle [listTitle]`
: Title of the list to retrieve items from. Specify `listId` or `listTitle` but not both

`--title [title]`
: (deprecated. Use `listTitle` instead) Title of the list to retrieve items from. Specify `id` or `title` but not both

`-q, --camlQuery [camlQuery]`
: CAML query to use to query the list of items with

`-f, --fields [fields]`
: Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested. Specify `camlQuery` or `fields` but not both

`-l, --filter [filter]`
: OData filter to use to query the list of items with. Specify `camlQuery` or `filter` but not both

`-p, --pageSize [pageSize]`
: Number of list items to return. Specify `camlQuery` or `pageSize` but not both

`-n, --pageNumber [pageNumber]`
: Page number to return if `pageSize` is specified (first page is indexed as value of 0)

--8<-- "docs/cmd/_global.md"

## Remarks

`pageNumber` is specified as a 0-based index. A value of `2` returns the third page of items.

If you want to specify a lookup type in the `properties` option, define which columns from the related list should be returned.

## Examples

Get all items from a list named Demo List

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x
```

From a list named _Demo List_ get all items with title _Demo list item_ using a CAML query

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --camlQuery "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo list item</Value></Eq></Where></Query></View>"
```

Get all items from a list with ID _935c13a0-cc53-4103-8b48-c1d0828eaa7f_

```sh
m365 spo listitem list --listId 935c13a0-cc53-4103-8b48-c1d0828eaa7f --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get all items from list named _Demo List_. For each item, retrieve the value of the _ID_, _Title_ and _Modified_ fields

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --fields "ID,Title,Modified"
```

Get all items from list named _Demo List_. For each item, retrieve the value of the _ID_, _Title_, _Modified_ fields, and the value of lookup field _Company_

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --fields "ID,Title,Modified,Company/Title"
```

From a list named _Demo List_ get all items with title _Demo list item_ using an OData filter

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --filter "Title eq 'Demo list item'"
```

From a list named _Demo List_ get the second batch of 10 items

```sh
m365 spo listitem list --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --pageSize 10 --pageNumber 2
```
