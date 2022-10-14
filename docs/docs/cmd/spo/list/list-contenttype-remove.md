# spo list contenttype remove

Removes content type from list

## Usage

```sh
m365 spo list contenttype remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

`-i, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-c, --contentTypeId <contentTypeId>`
: ID of the content type

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples

Remove content type with a specific id from the list retrieved by id in a specific site

```sh
m365 spo list contenttype remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --contentTypeId 0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A
```

Remove content type with a specific id from the list retrieved by title in a specific site

```sh
m365 spo list contenttype remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --contentTypeId 0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A
```

Remove content type with a specific id from the list retrieved by server relative URL in a specific site. This will not prompt for confirmation.

```sh
m365 spo list contenttype remove --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'sites/project-x/Documents' --contentTypeId 0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A --confirm
```
