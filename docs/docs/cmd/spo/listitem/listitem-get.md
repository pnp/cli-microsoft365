# spo listitem get

Gets a list item from the specified list

## Usage

```sh
m365 spo listitem get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved

`-i, --id <id>`
: ID of the item to retrieve.

`-l, --listId [listId]`
: ID of the list from which to retrieve the item. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list from which to retrieve the item. Specify `listId` or `listTitle` but not both

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested

--8<-- "docs/cmd/_global.md"

## Remarks

If you want to specify a lookup type in the `properties` option, define which columns from the related list should be returned.

## Examples

Get an item with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get an items _Title_ and _Created_ column with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created"
```

Get an items _Title_, _Created_ column and lookup column _Company_ with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Created,Company/Title"
```
