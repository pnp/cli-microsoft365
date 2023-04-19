# outlook message get

Retrieves specified message

## Usage

```sh
m365 outlook message get [options]
```

## Options

`-i, --id <id>`
: ID of the message

`--userId [userId]`
: ID of the user from which to retrieve the message. Specify either `userId` or `userPrincipalName`, but not both. This option is required when using application permissions.

`--userPrincipalName [userPrincipalName]`
: UPN of the user from which to retrieve the message Specify either `userId` or `userPrincipalName`, but not both. This option is required when using application permissions.

--8<-- "docs/cmd/_global.md"

## Examples

Get a specific message using delegated permissions

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA=
```

Get a specific message using delegated permissions from a shared mailbox

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA= --userPrincipalName sharedmailbox@tenant.com
```

Get a specific message from a specific user retrieved by user ID using application permissions

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA= --userId 6799fd1a-723b-4eb7-8e52-41ae530274ca
```

Get a specific message from a specific user retrieved by user principal name using application permissions

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA= --userPrincipalName user@tenant.com
```

## Response

=== "JSON"

    ```json
    {
      "id": "AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP_VgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA=",
      "createdDateTime": "2023-01-26T19:22:44Z",
      "lastModifiedDateTime": "2023-01-26T19:22:46Z",
      "changeKey": "CQAAABYAAAAhebtol4HnTZCmNsr9Gnh6AAPehsHb",
      "categories": [],
      "receivedDateTime": "2023-01-26T19:22:45Z",
      "sentDateTime": "2023-01-26T19:22:42Z",
      "hasAttachments": true,
      "internetMessageId": "<HE1P190MB032953D4D9C86FCEF5FFA8C4CECF9@HE1P190MB0329.EURP190.PROD.OUTLOOK.COM>",
      "subject": "Lorem ipsum",
      "bodyPreview": "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra.",
      "importance": "normal",
      "parentFolderId": "AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAuAAAAAABn1FpEFqPeR7YAnkzP_VgXAQAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAA=",
      "conversationId": "AAQkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAQAJfvGq77YHxJvRN73_QiuVw=",
      "conversationIndex": "AQHZMbuNl+8arvtgfEm9E3vf5CK5XA==",
      "isDeliveryReceiptRequested": false,
      "isReadReceiptRequested": false,
      "isRead": false,
      "isDraft": false,
      "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP%2BVgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
      "inferenceClassification": "focused",
      "body": {
        "contentType": "html",
        "content": "<html><head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><style type=\"text/css\" style=\"display:none\">\r\\\n<!--\r\\\np\r\\\n\t{margin-top:0;\r\\\n\tmargin-bottom:0}\r\\\n-->\r\\\n</style></head><body dir=\"ltr\"><div class=\"elementToProof ContentPasted0\" style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255)\">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra. Suspendisse cursus turpis vel urna volutpat congue. Etiam auctor nec nulla sed suscipit. Vestibulum rhoncus quis mi ac faucibus. Curabitur eget eleifend felis. Vestibulum ut dolor non elit molestie ornare. <br></div></body></html>"
      },
      "sender": {
        "emailAddress": {
          "name": "John Doe",
          "address": "john.doe@contoso.com"
        }
      },
      "from": {
        "emailAddress": {
          "name": "John Doe",
          "address": "john.doe@contoso.com"
        }
      },
      "toRecipients": [
        {
          "emailAddress": {
            "name": "Megan Bowen",
            "address": "megan.bowen@contoso.com"
          }
        }
      ],
      "ccRecipients": [],
      "bccRecipients": [],
      "replyTo": [],
      "flag": {
        "flagStatus": "notFlagged"
      }
    }
    ```

=== "Text"

    ```txt
    bccRecipients             : []
    body                      : {"contentType":"html","content":"<html><head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><style type=\"text/css\" style=\"display:none\">\r\n<!--\r\np\r\n\t{margin-top:0;\r\n\tmargin-bottom:0}\r\n-->\r\n</style></head><body dir=\"ltr\"><div class=\"elementToProof ContentPasted0\" style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255)\">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra. Suspendisse cursus turpis vel urna volutpat congue. Etiam auctor nec nulla sed suscipit. Vestibulum rhoncus quis mi ac faucibus. Curabitur eget eleifend felis. Vestibulum ut dolor non elit molestie ornare. <br></div></body></html>"}
    bodyPreview               : Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra.
    categories                : []
    ccRecipients              : []
    changeKey                 : CQAAABYAAAAhebtol4HnTZCmNsr9Gnh6AAPehsHb
    conversationId            : AAQkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAQAJfvGq77YHxJvRN73_QiuVw=
    conversationIndex         : AQHZMbuNl+8arvtgfEm9E3vf5CK5XA==
    createdDateTime           : 2023-01-26T19:22:44Z
    flag                      : {"flagStatus":"notFlagged"}
    from                      : {"emailAddress":{"name":"John Doe","address":"john.doe@contoso.com"}}
    hasAttachments            : true
    id                        : AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP_VgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA=
    importance                : normal
    inferenceClassification   : focused
    internetMessageId         : <HE1P190MB032953D4D9C86FCEF5FFA8C4CECF9@HE1P190MB0329.EURP190.PROD.OUTLOOK.COM>
    isDeliveryReceiptRequested: false
    isDraft                   : false
    isRead                    : false
    isReadReceiptRequested    : false
    lastModifiedDateTime      : 2023-01-26T19:22:46Z
    parentFolderId            : AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAuAAAAAABn1FpEFqPeR7YAnkzP_VgXAQAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAA=
    receivedDateTime          : 2023-01-26T19:22:45Z
    replyTo                   : []
    sender                    : {"emailAddress":{"name":"John Doe","address":"john.doe@contoso.com"}}
    sentDateTime              : 2023-01-26T19:22:42Z
    subject                   : Lorem ipsum
    toRecipients              : [{"emailAddress":{"name":"Megan Bowen","address":"megan.bowen@contoso.com"}}]
    webLink                   : https://outlook.office365.com/owa/?ItemID=AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP%2BVgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA%3D&exvsurl=1&viewmodel=ReadMessageItem
    ```

=== "CSV"

    ```csv
    id,createdDateTime,lastModifiedDateTime,changeKey,categories,receivedDateTime,sentDateTime,hasAttachments,internetMessageId,subject,bodyPreview,importance,parentFolderId,conversationId,conversationIndex,isDeliveryReceiptRequested,isReadReceiptRequested,isRead,isDraft,webLink,inferenceClassification,body,sender,from,toRecipients,ccRecipients,bccRecipients,replyTo,flag
    AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP_VgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA=,2023-01-26T19:22:44Z,2023-01-26T19:22:46Z,CQAAABYAAAAhebtol4HnTZCmNsr9Gnh6AAPehsHb,[],2023-01-26T19:22:45Z,2023-01-26T19:22:42Z,1,<HE1P190MB032953D4D9C86FCEF5FFA8C4CECF9@HE1P190MB0329.EURP190.PROD.OUTLOOK.COM>,Lorem ipsum,"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra.",normal,AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAuAAAAAABn1FpEFqPeR7YAnkzP_VgXAQAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAA=,AAQkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAQAJfvGq77YHxJvRN73_QiuVw=,AQHZMbuNl+8arvtgfEm9E3vf5CK5XA==,,,,,https://outlook.office365.com/owa/?ItemID=AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP%2BVgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA%3D&exvsurl=1&viewmodel=ReadMessageItem,focused,"{""contentType"":""html"",""content"":""<html><head>\r\n<meta http-equiv=\""Content-Type\"" content=\""text/html; charset=utf-8\""><style type=\""text/css\"" style=\""display:none\"">\r\n<!--\r\np\r\n\t{margin-top:0;\r\n\tmargin-bottom:0}\r\n-->\r\n</style></head><body dir=\""ltr\""><div class=\""elementToProof ContentPasted0\"" style=\""font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255)\"">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra. Suspendisse cursus turpis vel urna volutpat congue. Etiam auctor nec nulla sed suscipit. Vestibulum rhoncus quis mi ac faucibus. Curabitur eget eleifend felis. Vestibulum ut dolor non elit molestie ornare. <br></div></body></html>""}","{""emailAddress"":{""name"":""John Doe"",""address"":""john.doe@contoso.com""}}","{""emailAddress"":{""name"":""John Doe"",""address"":""john.doe@contoso.com""}}","[{""emailAddress"":{""name"":""Megan Bowen"",""address"":""megan.bowen@contoso.com""}}]",[],[],[],"{""flagStatus"":""notFlagged""}"
    ```

=== "Markdown"

    ```md
    # outlook message get --id "AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP_VgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA="

    Date: 4/2/2023

    ## AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP_VgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA=

    Property | Value
    ---------|-------
    id | AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP_VgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA=
    createdDateTime | 2023-01-26T19:22:44Z
    lastModifiedDateTime | 2023-01-26T19:22:46Z
    changeKey | CQAAABYAAAAhebtol4HnTZCmNsr9Gnh6AAPk7Plc
    categories | []
    receivedDateTime | 2023-01-26T19:22:45Z
    sentDateTime | 2023-01-26T19:22:42Z
    hasAttachments | true
    internetMessageId | <HE1P190MB032953D4D9C86FCEF5FFA8C4CECF9@HE1P190MB0329.EURP190.PROD.OUTLOOK.COM>
    subject | Lorem ipsum
    bodyPreview | Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra.
    importance | normal
    parentFolderId | AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAuAAAAAABn1FpEFqPeR7YAnkzP\_VgXAQAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAA=
    conversationId | AAQkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQAQAJfvGq77YHxJvRN73\_QiuVw=
    conversationIndex | AQHZMbuNl+8arvtgfEm9E3vf5CK5XA==
    isDeliveryReceiptRequested | false
    isReadReceiptRequested | false
    isRead | true
    isDraft | false
    webLink | https://outlook.office365.com/owa/?ItemID=AAMkAGUzZWVmZWU4LTM5ZmItNDA4My04OTIzLWY1MGUxNzdiYTQ0MQBGAAAAAABn1FpEFqPeR7YAnkzP%2BVgXBwAhebtol4HnTZCmNsr9Gnh6AAAAAAEMAAAhebtol4HnTZCmNsr9Gnh6AAPfHbtVAAA%3D&exvsurl=1&viewmodel=ReadMessageItem
    inferenceClassification | focused
    body | {"contentType":"html","content":"<html><head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><style type=\"text/css\" style=\"display:none\">\r\n<!--\r\np\r\n\t{margin-top:0;\r\n\tmargin-bottom:0}\r\n-->\r\n</style></head><body dir=\"ltr\"><div class=\"elementToProof ContentPasted0\" style=\"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255)\">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis vel diam gravida, auctor mauris nec, posuere tellus. Vivamus placerat, nunc ac cursus feugiat, arcu tellus mattis nisl, id cursus nisl lectus eu lacus. Praesent malesuada ut orci vitae viverra. Suspendisse cursus turpis vel urna volutpat congue. Etiam auctor nec nulla sed suscipit. Vestibulum rhoncus quis mi ac faucibus. Curabitur eget eleifend felis. Vestibulum ut dolor non elit molestie ornare. <br></div></body></html>"}
    sender | {"emailAddress":{"name":"John Doe","address":"john.doe@contoso.com"}}
    from | {"emailAddress":{"name":"John Doe","address":"john.doe@contoso.com"}}
    toRecipients | [{"emailAddress":{"name":"Megan Bowen","address":"megan.bowen@contoso.com"}}]
    ccRecipients | []
    bccRecipients | []
    replyTo | []
    flag | {"flagStatus":"notFlagged"}
    ```
