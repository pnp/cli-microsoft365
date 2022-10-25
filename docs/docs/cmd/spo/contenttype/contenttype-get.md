# spo contenttype get

Retrieves information about the specified list or site content type

## Usage

```sh
m365 spo contenttype get [options]
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

`-i, --id [id]`
: The ID of the content type to retrieve. Specify either id or name but not both

`-n, --name [name]`
: The name of the content type to retrieve. Specify either id or name but not both

--8<-- "docs/cmd/_global.md"

## Remarks

If no content type with the specified is found in the site or the list, you will get the _Content type with ID 0x010012 not found_ error.

## Examples

Retrieve site content type by id

```sh
m365 spo contenttype get --webUrl https://contoso.sharepoint.com/sites/contoso-sales --id 0x0100558D85B7216F6A489A499DB361E1AE2F
```

Retrieve site content type by name

```sh
m365 spo contenttype get --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'Document'
```

Retrieve list (retrieved by Title) content type 

```sh
m365 spo contenttype get --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --id 0x0100558D85B7216F6A489A499DB361E1AE2F
```

Retrieve list (retrieved by ID) content type 

```sh
m365 spo contenttype get --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listId '8c7a0fcd-9d64-4634-85ea-ce2b37b2ec0c' --id 0x0100558D85B7216F6A489A499DB361E1AE2F
```

Retrieve list (retrieved by URL) content type 

```sh
m365 spo contenttype get --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl '/Shared Documents' --id 0x0100558D85B7216F6A489A499DB361E1AE2F
```
