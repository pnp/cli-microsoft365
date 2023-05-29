# spo list sensitivitylabel ensure

Applies a default sensitivity label to the specified document library

## Usage

```sh
m365 spo list sensitivitylabel ensure [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the library is located.

`-n, --name <name>`
: The name of the label.

`-t, --listTitle [listTitle]`
: The title of the library on which to apply the label. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

`-l, --listId [listId]`
: The ID of the library on which to apply the label. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

`--listUrl [listUrl]`
: Server- or web-relative URL of the library on which to apply the label. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reaches general availability.

## Examples

Applies a sensitivity label to a document library based on the list title.

```sh
m365 spo list sensitivitylabel ensure --webUrl 'https://contoso.sharepoint.com' --listTitle 'Shared Documents' --name 'Confidential'
```

Applies a sensitivity label to a document library based on the list url.

```sh
m365 spo list sensitivitylabel ensure --webUrl 'https://contoso.sharepoint.com' --listUrl '/Shared Documents' --name 'Confidential'
```

Applies a sensitivity label to a document library based on the list id.

```sh
m365 spo list sensitivitylabel ensure --webUrl 'https://contoso.sharepoint.com' --listId 'b4cfa0d9-b3d7-49ae-a0f0-f14ffdd005f7' --name 'Confidential'
```

## Response

The command won't return a response on success.
