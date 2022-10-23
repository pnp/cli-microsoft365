# spo contenttype field remove

Removes a column from a site- or list content type

## Usage

```sh
m365 spo contenttype field remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the content type is located

`-l, --listTitle [listTitle]`
: Title of the list (if it is a list content type). Specify either `listTitle`, `listId` or `listUrl`.

`--listId [listId]`
: ID of the list (if it is a list content type). Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list (if it is a list content type). Specify either `listTitle`, `listId` or `listUrl`.

`-i, --contentTypeId <contentTypeId>`
: The ID of the content type to remove the column from

`-f, --fieldLinkId <fieldLinkId>`
: The ID of the column to remove

`-c, --updateChildContentTypes`
: Update child content types

`--confirm`
: Don't prompt for confirming removal of a column from content type

--8<-- "docs/cmd/_global.md"

## Examples

Remove column with a specific ID from a content type with a specific ID from a specific site.

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --confirm
```

Remove column with a specific ID from a content type with a specific ID from a specific site and updating the child content types

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --updateChildContentTypes
```

Remove fieldLink with a specific ID from a list (obtained by Title) content type with a specific ID on a spefici site.

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A60062F089A38C867747942DB2C3FC50FF6A" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --listTitle "Documents"
```

Remove fieldLink with a specific ID from a list (obtained by ID) content type with a specific ID on a spefici site.

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A60062F089A38C867747942DB2C3FC50FF6A" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --listId "8c7a0fcd-9d64-4634-85ea-ce2b37b2ec0c"
```

Remove fieldLink with a specific ID from a list (obtained by URL) content type with a specific ID on a spefici site.

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A60062F089A38C867747942DB2C3FC50FF6A" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --listUrl "/shared documents"
```
