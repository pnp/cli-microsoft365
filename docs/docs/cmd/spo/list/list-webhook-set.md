# spo list webhook set

Updates the specified webhook

## Usage

```sh
m365 spo list webhook set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located.

`-l, --listId [listId]`
: ID of the list. Specify either `listId`, `listTitle` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listId`, `listTitle` or `listUrl`.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listId`, `listTitle` or `listUrl`.

`-i, --id <id>`
: ID of the webhook to update.

`-n, --notificationUrl [notificationUrl]`
: The new notification url.

`-e, --expirationDateTime [expirationDateTime]`
: The new expiration date. _Note: Expiration Time cannot be more than 6 months in future_

`-c, --clientState [clientState]`
: A client state information that will be passed through notifications.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `id` doesn't refer to an existing webhook, you will get a `404 - "404 FILE NOT FOUND"` error.

## Examples

Update the notification url of a webhook with a specific ID attached to a list with a specific ID in a specific site to a specific URL

```sh
m365 spo list webhook set --webUrl https://contoso.sharepoint.com/sites/ninja --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id cc27a922-8224-4296-90a5-ebbc54da2e81 --notificationUrl https://contoso-functions.azurewebsites.net/webhook
```

Update the expiration date of a webhook with a specific ID attached to a list with a specific title in a specific site to a specfic date

```sh
m365 spo list webhook set --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e81 --expirationDateTime 2018-10-09T18:15
```

Update the notification url and clientState of a webhook with a specific ID attached to a list with a specific URL in a specific site to a specific URL and the expiration date to a specific date

```sh
m365 spo list webhook set --webUrl https://contoso.sharepoint.com/sites/ninja --listUrl '/sites/ninja/Documents' --id cc27a922-8224-4296-90a5-ebbc54da2e81 --notificationUrl https://contoso-functions.azurewebsites.net/webhook --expirationDateTime 2019-03-02 --clientState 'pnp-js-core-subscription'
```

## Respone

The command won't return a response on success.
