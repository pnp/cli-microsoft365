# teams message reply list

Retrieves replies to a message from a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message reply list [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel that contains the message

`-m, --messageId <messageId>`
: The ID of the message to retrieve replies for

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve replies to a message from a Microsoft Teams team if you are a member of that team.

## Examples

Retrieve the replies from a specified message from a channel of the Microsoft Teams team

```sh
m365 teams message reply list --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --messageId 1540747442203
```

## Response

=== "JSON"

    ``` json
    [
      {
        "id": "1666799841244",
        "replyToId": "1666799520731",
        "etag": "1666799841244",
        "messageType": "message",
        "createdDateTime": "2022-10-26T15:57:21.244Z",
        "lastModifiedDateTime": "2022-10-26T15:57:21.244Z",
        "lastEditedDateTime": null,
        "deletedDateTime": null,
        "subject": null,
        "summary": null,
        "chatId": null,
        "importance": "normal",
        "locale": "en-us",
        "webUrl": "https://teams.microsoft.com/l/message/19%3Ae2916df2b11046beba42d22da898383f%40thread.tacv2/1666799841244?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022&createdTime=1666799841244&parentMessageId=1666799520731",
        "policyViolation": null,
        "eventDetail": null,
        "from": {
          "application": null,
          "device": null,
          "user": {
            "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
            "displayName": "Nico De Cleyre",
            "userIdentityType": "aadUser",
            "tenantId": "0cac6cda-2e04-4a3d-9c16-9c91470d7022"
          }
        },
        "body": {
          "contentType": "html",
          "content": "<div><div>second reply <span class=\"animated-emoticon-20\" title=\"Grinning face with smiling eyes\" type=\"grinningfacewithsmilingeyes\"><img itemid=\"grinningfacewithsmilingeyes\" itemscope=\"\" itemtype=\"http://schema.skype.com/Emoji\" src=\"https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/grinningfacewithsmilingeyes/default/20_f.png?v=v8\" alt=\"😄\" style=\"width:20px; height:20px\"></span></div>\\\n\n\\\n</div>"
        },
        "channelIdentity": {
          "teamId": "aee5a2c9-b1df-45ac-9964-c708e760a045",
          "channelId": "19:e2916df2b11046beba42d22da898383f@thread.tacv2"
        },
        "attachments": [],
        "mentions": [],
        "reactions": []
      },
      {
        "id": "1666799833162",
        "replyToId": "1666799520731",
        "etag": "1666799833162",
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
        "webUrl": "https://teams.microsoft.com/l/message/19%3Ae2916df2b11046beba42d22da898383f%40thread.tacv2/1666799833162?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022&createdTime=1666799833162&parentMessageId=1666799520731",
        "policyViolation": null,
        "eventDetail": null,
        "from": {
          "application": null,
          "device": null,
          "user": {
            "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
            "displayName": "Nico De Cleyre",
            "userIdentityType": "aadUser",
            "tenantId": "0cac6cda-2e04-4a3d-9c16-9c91470d7022"
          }
        },
        "body": {
          "contentType": "text",
          "content": "First reply"
        },
        "channelIdentity": {
          "teamId": "aee5a2c9-b1df-45ac-9964-c708e760a045",
          "channelId": "19:e2916df2b11046beba42d22da898383f@thread.tacv2"
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
    -------------  --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    1666799841244  <div><div>second reply <span class="animated-emoticon-20" title="Grinning face with smiling eyes" type="grinningfacewithsmilingeyes"><img itemid="grinningfacewithsmilingeyes" itemscope="" itemtype="http://schema.skype.com/Emoji" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/grinningfacewithsmilingeyes/default/20_f.png?v=v8" alt="😄" style="width:20px; height:20px"></span></div>


    </div>
    1666799833162  First reply                                                                                                                                                                                                                                                                                                                              

    ```

=== "CSV"

    ``` text
    id,body
    1666799841244,"<div><div>second reply <span class=""animated-emoticon-20"" title=""Grinning face with smiling eyes"" type=""grinningfacewithsmilingeyes""><img itemid=""grinningfacewithsmilingeyes"" itemscope="""" itemtype=""http://schema.skype.com/Emoji"" src=""https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/grinningfacewithsmilingeyes/default/20_f.png?v=v8"" alt=""😄"" style=""width:20px; height:20px""></span></div>


    </div>"
    1666799833162,First reply

        
    ```
