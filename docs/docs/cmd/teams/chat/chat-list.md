# teams chat list

Lists all Microsoft Teams chat conversations for the current or a specific user.

## Usage

```sh
m365 teams chat list [options]
```

## Options

`-t, --type [type]`
: The chat type to optionally filter chat conversations by type. The value can be `oneOnOne`, `group` or `meeting`.

`--userId [userId]`
: ID of the user. Has to be specified when using application permissions. Specify either `userId` or `userName`, but not both.

`--userName [userName]`
: UPN of the user. Has to be specified when using application permissions. Specify either `userId` or `userName`, but not both.

--8<-- "docs/cmd/_global.md"

## Examples

List all the Microsoft Teams chat conversations of the current user.

```sh
m365 teams chat list
```

List only the one on one Microsoft Teams chat conversations of a specific user retrieved by id.

```sh
m365 teams chat list --userId e6296ed0-4b7d-4ace-aed4-f6b7371ce060 --type oneOnOne
```

List only the group Microsoft Teams chat conversations of a specific user retrieved by mail

```sh
m365 teams chat list --userName 'john@contoso.com' --type group 
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
        "topic": null,
        "createdDateTime": "2022-11-05T13:06:25.218Z",
        "lastUpdatedDateTime": "2022-11-05T13:06:25.218Z",
        "chatType": "oneOnOne",
        "webUrl": "https://teams.microsoft.com/l/chat/19%3A2da4c29f6d7041eca70b638b43d45437%40thread.v2/0?tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b",
        "tenantId": "446355e4-e7e3-43d5-82f8-d7ad8272d55b",
        "onlineMeetingInfo": null,
        "viewpoint": {
          "isHidden": false,
          "lastMessageReadDateTime": "2022-11-05T13:06:30.582Z"
        }
      }
    ]
    ```

=== "Text"

    ```text
    id                                             topic  chatType
    ---------------------------------------------  -----  --------
    19:2da4c29f6d7041eca70b638b43d45437@thread.v2  null   oneOnOne
    ```

=== "CSV"

    ```csv
    id,topic,chatType
    19:2da4c29f6d7041eca70b638b43d45437@thread.v2,,oneOnOne
    ```

==="Markdown"

    ```md

# teams chat list --type "group"

Date: 5/8/2023

## 19:b4874b8c8472461d9ac7b033ca782b33@thread.v2

Property | Value
---------|-------
id | 19:b4874b8c8472461d9ac7b033ca782b33@thread.v2
topic | Just a conversation
createdDateTime | 2021-03-04T18:18:50.303Z
lastUpdatedDateTime | 2023-05-07T18:15:20.291Z
chatType | group
webUrl | https://teams.microsoft.com/l/chat/19%3Ab4874b8c8472461d9ac7b033ca782b33%40thread.v2/0?tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4
tenantId | de348bc7-1aeb-4406-8cb3-97db021cadb4
onlineMeetingInfo | null
viewpoint | {"isHidden":false,"lastMessageReadDateTime":"2023-05-07T18:15:20.309Z"}

      ```
