# graph subscription add

Creates a Microsoft Graph subscription

## Usage

```sh
m365 graph subscription add [options]
```

## Options

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

--8<-- "docs/cmd/_global.md"

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

## Response

=== "JSON"

```json
{
  "id": "3eceb2b5-9bb0-41da-a931-a919b8e8e553",
  "resource": "groups",
  "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e",
  "changeType": "updated",
  "clientState": null,
  "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
  "notificationQueryOptions": null,
  "lifecycleNotificationUrl": null,
  "expirationDateTime": "2022-10-31T15:08:23.461Z",
  "creatorId": "411edae6-e8e6-4dbd-9a02-2cb6e319aa08",
  "includeResourceData": null,
  "latestSupportedTlsVersion": "v1_2",
  "encryptionCertificate": null,
  "encryptionCertificateId": null,
  "notificationUrlAppId": null
}
```

=== "Text"

    ``` text
applicationId            : 31359c7f-bd7e-475c-86db-fdb8c937548e
changeType               : updated
clientState              : null
creatorId                : 411edae6-e8e6-4dbd-9a02-2cb6e319aa08
encryptionCertificate    : null
encryptionCertificateId  : null
expirationDateTime       : 2022-10-31T15:09:15.356Z
id                       : 094aeced-1f16-44ff-a4c8-3c0610b824a0
includeResourceData      : null
latestSupportedTlsVersion: v1_2
lifecycleNotificationUrl : null
notificationQueryOptions : null
notificationUrl          : https://webhook.azurewebsites.net/api/send/myNotifyClient
notificationUrlAppId     : null
resource                 : groups    
````

=== "CSV"

    ``` text
id,resource,applicationId,changeType,clientState,notificationUrl,notificationQueryOptions,lifecycleNotificationUrl,expirationDateTime,creatorId,includeResourceData,latestSupportedTlsVersion,encryptionCertificate,encryptionCertificateId,notificationUrlAppId
e926b017-fc99-41d8-b9cf-3d2f8663e2fa,groups,31359c7f-bd7e-475c-86db-fdb8c937548e,updated,,https://webhook.azurewebsites.net/api/send/myNotifyClient,,,2022-10-31T15:09:41.241Z,411edae6-e8e6-4dbd-9a02-2cb6e319aa08,,v1_2,,,
````

