# teams chat message list

Lists all messages from a Microsoft Teams chat conversation.

## Usage

```sh
m365 teams chat message list [options]
```

## Options

`-i, --chatId <chatId>`
: The ID of the chat conversation

--8<-- "docs/cmd/_global.md"

## Examples

List the messages from a Microsoft Teams chat conversation

```sh
m365 teams chat message list --chatId 19:2da4c29f6d7041eca70b638b43d45437@thread.v2
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "1667653590582",
        "replyToId": null,
        "etag": "1667653590582",
        "messageType": "message",
        "createdDateTime": "2022-11-05T13:06:30.582Z",
        "lastModifiedDateTime": "2022-11-05T13:06:30.582Z",
        "lastEditedDateTime": null,
        "deletedDateTime": null,
        "subject": null,
        "summary": null,
        "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
        "importance": "normal",
        "locale": "en-us",
        "webUrl": null,
        "channelIdentity": null,
        "policyViolation": null,
        "eventDetail": null,
        "from": {
          "application": null,
          "device": null,
          "user": {
            "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
            "displayName": "John Doe",
            "userIdentityType": "aadUser",
            "tenantId": "446355e4-e7e3-43d5-82f8-d7ad8272d55b"
          }
        },
        "body": {
          "contentType": "html",
          "content": "<p>Hello world</p>"
        },
        "attachments": [],
        "mentions": [],
        "reactions": []
      }
    ]
    ```

=== "Text"

    ```text
    id             shortBody
    -------------  -------------------------
    1667653590582  <p>Hello world</p>
    ```

=== "CSV"

    ```csv
    id,shortBody
    1667653590582,<p>Hello world</p>
    ```

==="Markdown"

    ```md
# teams chat message list --chatId "19:04355ecd-2124-4097-bc2b-c2295a71d7a3_e1251b10-1ba4-49e3-b35a-933e3f21772b@unq.gbl.spaces"

Date: 5/8/2023

## 1662642685689

Property | Value
---------|-------
id | 1662642685689
replyToId | null
etag | 1662642685689
messageType | message
createdDateTime | 2022-09-08T13:11:25.689Z
lastModifiedDateTime | 2022-09-08T13:11:25.689Z
lastEditedDateTime | null
deletedDateTime | null
subject | null
summary | null
chatId | 19:04355ecd-2124-4097-bc2b-c2295a71d7a3\_e1251b10-1ba4-49e3-b35a-933e3f21772b@unq.gbl.spaces
importance | normal
locale | en-us
webUrl | null
channelIdentity | null
policyViolation | null
eventDetail | null
from | {"application":null,"device":null,"user":{"id":"78ccf530-bbf0-47e4-aae6-da5f8c6fb142","displayName":"John Doe","userIdentityType":"aadUser","tenantId":"446355e4-e7e3-43d5-82f8-d7ad8272d55b"}}
body | <attachment id="ead8f1e253584d289d760264f59c9e61"></attachment>
attachments | [{"id":"ead8f1e253584d289d760264f59c9e61","contentType":"application/vnd.microsoft.card.adaptive","contentUrl":null,"content":"{\r\n  \"type\": \"AdaptiveCard\",\r\n  \"body\": [\r\n    {\r\n      \"color\": \"accent\",\r\n      \"horizontalAlignment\": \"center\",\r\n      \"size\": \"extraLarge\",\r\n      \"text\": \"Reporting of inappropriate content\",\r\n      \"weight\": \"bolder\",\r\n      \"id\": \"Title\",\r\n      \"spacing\": \"Medium\",\r\n      \"type\": \"TextBlock\"\r\n    },\r\n    {\r\n      \"text\": \"I would like to bring your attention to below inappropriate content.\",\r\n      \"wrap\": true,\r\n      \"spacing\": \"Medium\",\r\n      \"type\": \"TextBlock\"\r\n    },\r\n    {\r\n      \"columns\": [\r\n        {\r\n          \"width\": \"stretch\",\r\n          \"items\": [\r\n            {\r\n              \"text\": \"Content reported by\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"Content created by\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"Content\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            }\r\n          ],\r\n          \"type\": \"Column\"\r\n        },\r\n        {\r\n          \"width\": \"stretch\",\r\n          \"items\": [\r\n            {\r\n              \"text\": \"John Doe\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"Amy Jones\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"\\n\\n\\nThat was the scrap, ugly, and foolish thing to do. Totally dumb\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            }\r\n          ],\r\n          \"type\": \"Column\"\r\n        }\r\n      ],\r\n      \"type\": \"ColumnSet\"\r\n    },\r\n    {\r\n      \"color\": \"accent\",\r\n      \"size\": \"medium\",\r\n      \"text\": \"Content Moderator Ratings\",\r\n      \"weight\": \"bolder\",\r\n      \"wrap\": true,\r\n      \"spacing\": \"Large\",\r\n      \"type\": \"TextBlock\"\r\n    },\r\n    {\r\n      \"text\": \"Below are the ratings reported by the Content moderator.\",\r\n      \"wrap\": true,\r\n      \"type\": \"TextBlock\"\r\n    },\r\n    {\r\n      \"columns\": [\r\n        {\r\n          \"width\": \"stretch\",\r\n          \"items\": [\r\n            {\r\n              \"text\": \"Category1 Score\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"Category2 Score\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"Category3 Score\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"Review Recommended?\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"Detected profanity terms\",\r\n              \"weight\": \"bolder\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            }\r\n          ],\r\n          \"type\": \"Column\"\r\n        },\r\n        {\r\n          \"width\": \"stretch\",\r\n          \"items\": [\r\n            {\r\n              \"text\": \"0.000200446273083799\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"0.151560887694359\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"0.987999975681305\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"True\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            },\r\n            {\r\n              \"text\": \"scrap, ugly, foolish thing . Totally dumb\\n\\n\",\r\n              \"wrap\": true,\r\n              \"type\": \"TextBlock\"\r\n            }\r\n          ],\r\n          \"type\": \"Column\"\r\n        }\r\n      ],\r\n      \"type\": \"ColumnSet\"\r\n    },\r\n    {\r\n      \"text\": \"Lets make our workplace better every day!\",\r\n      \"weight\": \"lighter\",\r\n      \"wrap\": true,\r\n      \"spacing\": \"Large\",\r\n      \"type\": \"TextBlock\"\r\n    }\r\n  ],\r\n  \"actions\": [\r\n    {\r\n      \"url\": \"https://teams.microsoft.com/l/message/19:UP\_ykQW1D9eML0XD3kul-wRSYF36BZS94SI7CRFvHQY1@thread.tacv2/1662639570789\",\r\n      \"title\": \"Open message\",\r\n      \"type\": \"Action.OpenUrl\"\r\n    }\r\n  ],\r\n  \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\",\r\n  \"version\": \"1.4\"\r\n}","name":null,"thumbnailUrl":null,"teamsAppId":null}]
mentions | []
reactions | []
shortBody | <attachment id="ead8f1e253584d289d760264f59c9e61">...
      ```
