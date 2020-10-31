# graph subscription add

Creates a Microsoft Graph subscription

## Usage

```sh
m365 graph subscription add [options]
```

## Options

`-h, --help`
: output usage information

`-r, --resource <resource>`
: The resource that will be monitored for changes

`-c, --changeType <changeType>`
: The type of change in the subscribed resource that will raise a notification. The supported values are: `created`, `updated`, `deleted`. Multiple values can be combined using a comma-separated list

`-u, --notificationUrl <notificationUrl>`
: The URL of the endpoint that will receive the notifications. This URL must use the HTTPS protocol

`-e, --expirationDateTime [expirationDateTime]`
: The date and time when the webhook subscription expires. The time is in UTC, and can be an amount of time from subscription creation that varies for the resource subscribed to. If not specified, the maximum allowed expiration for the specified resource will be used

`-s, --clientState [clientState]`
: The value of the clientState property sent by the service in each notification. The maximum length is 128 characters

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

On personal OneDrive, you can subscribe to the root folder or any subfolder in that drive. On OneDrive for Business, you can subscribe to only the root folder.

Notifications are sent for the requested types of changes on the subscribed folder, or any file, folder, or other `driveItem` instances in its hierarchy. You cannot subscribe to `drive` or `driveItem` instances that are not folders, such as individual files.

In Outlook, delegated permission supports subscribing to items in folders in only the signed-in user's mailbox.
That means, for example, you cannot use the delegated permission Calendars.Read to subscribe to events in another user’s mailbox.

To subscribe to change notifications of Outlook contacts, events, or messages in shared or delegated folders:

- Use the corresponding application permission to subscribe to changes of items in a folder or mailbox of any user in the tenant.
- Do not use the Outlook sharing permissions (Contacts.Read.Shared, Calendars.Read.Shared, Mail.Read.Shared, and their read/write counterparts), as they do not support subscribing to change notifications on items in shared or delegated folders.

## Examples

Create a subscription

```sh
m365 graph subscription add --resource "me/mailFolders('Inbox')/messages" --changeType "updated" --notificationUrl "https://webhook.azurewebsites.net/api/send/myNotifyClient" --expirationDateTime "2016-11-20T18:23:45.935Z" --clientState "secretClientState"

```

Create a subscription on multiple change types

```sh
m365 graph subscription add --resource groups --changeType updated,deleted --notificationUrl "https://webhook.azurewebsites.net/api/send/myNotifyClient" --expirationDateTime "2016-11-20T18:23:45.935Z" --clientState "secretClientState"

```

Create a subscription using the maximum allowed expiration for Group resources

```sh
m365 graph subscription add --resource groups --changeType "updated" --notificationUrl "https://webhook.azurewebsites.net/api/send/myNotifyClient"
```
