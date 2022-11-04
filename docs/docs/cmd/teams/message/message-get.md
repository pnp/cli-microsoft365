# teams message get

Retrieves a message from a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message get [options]
```

## Options

`-t, --teamId <teamId>`
: The ID of the team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel that contains the message

`-i, --id <id>`
: The ID of the message to retrieve

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve a message from a Microsoft Teams team if you are a member of that team.

## Examples

Retrieve the specified message from a channel of the Microsoft Teams team

```sh
m365 teams message get --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --id 1540747442203
```

## Response

=== "JSON"

    ``` json
    {
      "id": "1666799520731",
      "replyToId": null,
      "etag": "1666799582385",
      "messageType": "message",
      "createdDateTime": "2022-10-26T15:52:00.731Z",
      "lastModifiedDateTime": "2022-10-26T15:53:02.385Z",
      "lastEditedDateTime": "2022-10-26T15:53:02.288Z",
      "deletedDateTime": null,
      "subject": "Second message Title",
      "summary": null,
      "chatId": null,
      "importance": "normal",
      "locale": "en-us",
      "webUrl": "https://teams.microsoft.com/l/message/19%3Ae2916df2b11046beba42d22da898383f%40thread.tacv2/1666799520731?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba&createdTime=1666799520731&parentMessageId=1666799520731",
      "policyViolation": null,
      "eventDetail": null,
      "from": {
        "application": null,
        "device": null,
        "user": {
          "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
          "displayName": "Nico De Cleyre",
          "userIdentityType": "aadUser",
          "tenantId": "92e59666-257b-49c3-b1fa-1bae8107f6ba"
        }
      },
      "body": {
        "contentType": "text",
        "content": "second message!"
      },
      "channelIdentity": {
        "teamId": "aee5a2c9-b1df-45ac-9964-c708e760a045",
        "channelId": "19:e2916df2b11046beba42d22da898383f@thread.tacv2"
      },
      "attachments": [],
      "mentions": [],
      "reactions": []
    }
    ```

=== "Text"

    ``` text
    attachments         : []
    body                : {"contentType":"text","content":"second message!"}
    channelIdentity     : {"teamId":"aee5a2c9-b1df-45ac-9964-c708e760a045","channelId":"19:e2916df2b11046beba42d22da898383f@thread.tacv2"}
    chatId              : null
    createdDateTime     : 2022-10-26T15:52:00.731Z
    deletedDateTime     : null
    etag                : 1666799582385
    eventDetail         : null
    from                : {"application":null,"device":null,"user":{"id":"78ccf530-bbf0-47e4-aae6-da5f8c6fb142","displayName":"Nico De Cleyre","userIdentityType":"aadUser","tenantId":"92e59666-257b-49c3-b1fa-1bae8107f6ba"}}
    id                  : 1666799520731
    importance          : normal
    lastEditedDateTime  : 2022-10-26T15:53:02.288Z
    lastModifiedDateTime: 2022-10-26T15:53:02.385Z
    locale              : en-us
    mentions            : []
    messageType         : message
    policyViolation     : null
    reactions           : []
    replyToId           : null
    subject             : Second message Title
    summary             : null
    webUrl              : https://teams.microsoft.com/l/message/19%3Ae2916df2b11046beba42d22da898383f%40thread.tacv2/1666799520731?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba&createdTime=1666799520731&parentMessageId=1666799520731
    ```

=== "CSV"

    ``` text
    id,replyToId,etag,messageType,createdDateTime,lastModifiedDateTime,lastEditedDateTime,deletedDateTime,subject,summary,chatId,importance,locale,webUrl,policyViolation,eventDetail,from,body,channelIdentity,attachments,mentions,reactions
    1666799520731,,1666799582385,message,2022-10-26T15:52:00.731Z,2022-10-26T15:53:02.385Z,2022-10-26T15:53:02.288Z,,Second message Title,,,normal,en-us,https://teams.microsoft.com/l/message/19%3Ae2916df2b11046beba42d22da898383f%40thread.tacv2/1666799520731?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba&createdTime=1666799520731&parentMessageId=1666799520731,,,"{""application"":null,""device"":null,""user"":{""id"":""78ccf530-bbf0-47e4-aae6-da5f8c6fb142"",""displayName"":""Nico De Cleyre"",""userIdentityType"":""aadUser"",""tenantId"":""92e59666-257b-49c3-b1fa-1bae8107f6ba""}}","{""contentType"":""text"",""content"":""second message!""}","{""teamId"":""aee5a2c9-b1df-45ac-9964-c708e760a045"",""channelId"":""19:e2916df2b11046beba42d22da898383f@thread.tacv2""}",[],[],[]
    ```
