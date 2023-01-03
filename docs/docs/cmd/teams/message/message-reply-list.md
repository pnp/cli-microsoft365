# teams message reply list

Retrieves replies to a message from a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message reply list [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the team where the channel is located.

`-c, --channelId <channelId>`
: The ID of the channel that contains the message.

`-m, --messageId <messageId>`
: The ID of the message to retrieve replies for.

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve replies to a message from a Microsoft Teams team if you are a member of that team.

## Examples

Retrieve the replies from a specified message from a channel of the Microsoft Teams team.

```sh
m365 teams message reply list --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --messageId 1540747442203
```

## Response

=== "JSON"

    ``` json
    [
      {
        "id": "1540747442203",
        "replyToId": "1666799520731",
        "etag": "1540747442203",
        "messageType": "message",
        "createdDateTime": "2022-10-26T15:57:13.162Z",
        "lastModifiedDateTime": "2022-10-26T15:57:13.162Z",
        "lastEditedDateTime": null,
        "deletedDateTime": null,
        "subject": null,
        "summary": null,
        "chatId": null,
        "importance": "normal",
        "locale": "en-us",
        "webUrl": "https://teams.microsoft.com/l/message/19%388f7e66a8dfe42be92db19505ae912a8%40thread.tacv2/1540747442203?groupId=5f5d7b71-1161-44d8-bcc1-3da710eb4171&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba&createdTime=1540747442203&parentMessageId=1666799520731",
        "policyViolation": null,
        "eventDetail": null,
        "from": {
          "application": null,
          "device": null,
          "user": {
            "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
            "displayName": "John Doe",
            "userIdentityType": "aadUser",
            "tenantId": "92e59666-257b-49c3-b1fa-1bae8107f6ba"
          }
        },
        "body": {
          "contentType": "text",
          "content": "First reply"
        },
        "channelIdentity": {
          "teamId": "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
          "channelId": "19:88f7e66a8dfe42be92db19505ae912a8@thread.tacv2"
        },
        "attachments": [],
        "mentions": [],
        "reactions": []
      }
    ]
    ```

=== "Text"

    ``` text
    id             body
    -------------  -----------
    1540747442203  First reply
    ```

=== "CSV"

    ``` text
    id,body
    1540747442203,First reply
    ```

=== "Markdown"

    ```md
    # teams message reply list --teamId "5f5d7b71-1161-44d8-bcc1-3da710eb4171" --channelId "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype" --messageId "1540747442203"

    Date: 1/3/2023

    ## undefined (1672744052534)

    Property | Value
    ---------|-------
    id | 1540747442203
    replyToId | 1666799520731
    etag | 1540747442203
    messageType | message
    createdDateTime | 2022-10-26T15:57:13.162Z
    lastModifiedDateTime | 2022-10-26T15:57:13.162Z
    lastEditedDateTime | null
    deletedDateTime | null
    subject | null
    summary | null
    chatId | null
    importance | normal
    locale | en-us
    webUrl | https://teams.microsoft.com/l/message/19%388f7e66a8dfe42be92db19505ae912a8%40thread.tacv2/1540747442203?groupId=5f5d7b71-1161-44d8-bcc1-3da710eb4171&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba&createdTime=1540747442203&parentMessageId=1666799520731
    policyViolation | null
    eventDetail | null
    from | {"application":null,"device":null,"user":{"id":"78ccf530-bbf0-47e4-aae6-da5f8c6fb142","displayName":"John Doe","userIdentityType":"aadUser","tenantId":"92e59666-257b-49c3-b1fa-1bae8107f6ba"}}
    body | First reply
    channelIdentity | {"teamId":"5f5d7b71-1161-44d8-bcc1-3da710eb4171","channelId":"19:88f7e66a8dfe42be92db19505ae912a8@thread.skype"}
    attachments | []
    mentions | []
    reactions | []
    ```
