# spo list webhook add

Adds a new webhook to the specified list

## Usage

```sh
m365 spo list webhook add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located.

`-l, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-n, --notificationUrl <notificationUrl>`
: The notification URL.

`-e, --expirationDateTime [expirationDateTime]`
: The expiration date. Will be set to max (6 months from today) if not provided.

`-c, --clientState [clientState]`
: A client state information that will be passed through notifications.

--8<-- "docs/cmd/_global.md"

## Examples

Add a web hook to the list retrieved by Title located in a specific site with a specific notification url and the default expiration date

```sh
m365 spo list webhook add --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --notificationUrl https://contoso-funcions.azurewebsites.net/webhook
```

Add a web hook to the list retrieved by URL located in a specific site with a specific notification url and a specific expiration date

```sh
m365 spo list webhook add --webUrl https://contoso.sharepoint.com/sites/ninja --listUrl '/sites/ninja/Documents' --notificationUrl https://contoso-funcions.azurewebsites.net/webhook --expirationDateTime 2019-01-21
```

Add a web hook to the list retrieved by ID located in a specific site with a specific notification url, a specific expiration date and a client state


```sh
m365 spo list webhook add --webUrl https://contoso.sharepoint.com/sites/ninja --listId '3d6aefa0-f438-4789-b0cd-6e865f5d65b5' --notificationUrl https://contoso-funcions.azurewebsites.net/webhook --expirationDateTime '2019-03-02T18:15' --clientState "Hello State!"
```
