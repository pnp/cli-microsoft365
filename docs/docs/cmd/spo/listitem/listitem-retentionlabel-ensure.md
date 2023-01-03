# spo listitem retentionlabel ensure

Apply a retention label to a list item

## Usage

```sh
m365 spo listitem retentionlabel ensure [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the retentionlabel from a listitem to apply is located

`--listItemId <listItemId>`
: The ID of the list item for which the retention label should be applied.

`--listId [listId]`
: ID of the list where the retention label should be applied. Specify either `listTitle`, `listId` or `listUrl`

`--listTitle [listTitle]`
: Title of the list where the retention label should be applied. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`-n, --name [name]`
: The name of the retention label. Specify either `name` or `id`.

`-i, --id [id]`
: The id of the retention label. Specify either `name` or `id`.

--8<-- "docs/cmd/_global.md"

## Examples

Applies the retention label _Some label_ to a list item in a given site based on the list id and label name

```sh
m365 spo listitem retentionlabel ensure --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --listItemId 1 --name 'Some label'
```

Applies a retention label to a list item in a given site based on the list title and label id

```sh
m365 spo listitem retentionlabel ensure --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --listItemId 1 --id '7a621a91-063b-461b-aff6-d713d5fb23eb'
```

Applies the retention label _Some label_ to a list item in a given site based on the server relative list url

```sh
m365 spo listitem retentionlabel ensure --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl /sites/project-x/lists/TestList --listItemId 1 --name 'Some label'
```

## Response

The command won't return a response on success.
