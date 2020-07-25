# spo list webhook set

Updates the specified webhook

## Usage

```sh
m365 spo list webhook set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list which contains the webhook is located

`-l, --listId [listId]`
: ID of the list which contains the webhook which should be updated. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list which contains the webhook which should be updated. Specify either `listId` or `listTitle` but not both

`-i, --id [id]`
: ID of the webhook to update

`-n, --notificationUrl [notificationUrl]`
: The new notification url

`-e, --expirationDateTime [expirationDateTime]`
: The new expiration date

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified `id` doesn't refer to an existing webhook, you will get a `404 - "404 FILE NOT FOUND"` error.

## Examples

Update the notification url of a webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ which belongs to a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/ninja_ to _https://contoso-functions.azurewebsites.net/webhook_

```sh
m365 spo list webhook set --webUrl https://contoso.sharepoint.com/sites/ninja --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id cc27a922-8224-4296-90a5-ebbc54da2e81 --notificationUrl https://contoso-functions.azurewebsites.net/webhook
```

Update the expiration date of a webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ which belongs to a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/ninja_ to _October 9th, 2018 at 6:15 PM_

```sh
m365 spo list webhook set --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e81 --expirationDateTime 2018-10-09T18:15
```

From the webhook with ID _cc27a922-8224-4296-90a5-ebbc54da2e81_ which belongs to a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/ninja_ update the notification url to _https://contoso-functions.azurewebsites.net/webhook_
and the expiration date to _March 2nd, 2019_

```sh
m365 spo list webhook set --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --id cc27a922-8224-4296-90a5-ebbc54da2e81 --notificationUrl https://contoso-functions.azurewebsites.net/webhook --expirationDateTime 2019-03-02
```