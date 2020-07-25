# spo list webhook add

Adds a new webhook to the specified list

## Usage

```sh
m365 spo list webhook add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the list to add the webhook to is located

`-l, --listId [listId]`
: ID of the list to which the webhook which should be added. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list to which the webhook which should be added. Specify either `listId` or `listTitle` but not both

`-n, --notificationUrl <notificationUrl>`
: The notification url

`-e, --expirationDateTime [expirationDateTime]`
: The expiration date. Will be set to max (6 months from today) if not provided

`-c, --clientState [clientState]`
: A client state information that will be passed through notifications

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Add a web hook to the list Documents located in site _https://contoso.sharepoint.com/sites/ninja_ with the notification url
_https://contoso-funcions.azurewebsites.net/webhook_ and the default expiration
date

```sh
m365 spo list webhook add --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --notificationUrl https://contoso-funcions.azurewebsites.net/webhook
```

Add a web hook to the list Documents located in site _https://contoso.sharepoint.com/sites/ninja_ with the notification url
_https://contoso-funcions.azurewebsites.net/webhook_ and an expiration date of _January 21st, 2019_

```sh
m365 spo list webhook add --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --notificationUrl https://contoso-funcions.azurewebsites.net/webhook --expirationDateTime 2019-01-21
```

Add a web hook to the list Documents located in site _https://contoso.sharepoint.com/sites/ninja_ with the notification url
_https://contoso-funcions.azurewebsites.net/webhook_, a very specific expiration date of _6:15 PM on March 2nd, 2019_ and
a client state

```sh
m365 spo list webhook add --webUrl https://contoso.sharepoint.com/sites/ninja --listTitle Documents --notificationUrl https://contoso-funcions.azurewebsites.net/webhook --expirationDateTime '2019-03-02T18:15' --clientState "Hello State!"
```