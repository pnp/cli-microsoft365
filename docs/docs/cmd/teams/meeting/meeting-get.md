# teams meeting get

Get specified meeting details

## Usage

```sh
m365 teams meeting get [options]
```

## Options

`-u, --userId [userId]`
: The id of the user, omit to retrieve meetings for current signed in user. Use either `id`, `userName`, or `email`, but not multiple.

`-n, --userName [userName]`
: The name of the user, omit to retrieve meetings for current signed in user. Use either `id`, `userName`, or `email`, but not multiple.

`--email [email]`
: The email of the user, omit to retrieve meetings for current signed in user. Use either `id`, `userName`, or `email`, but not multiple.

`-j, --joinUrl <joinUrl>`
: The join URL from a calendar invite.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    When you connect with application permissions, you should [add an application access policy](https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy) for the command to work.

## Examples

Retrieve meeting details for given meeting with userId and joinurl

```sh
m365 teams meeting get --userId 00000000-0000-0000-0000-000000000000 --joinUrl 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_MGQ4MDQyNTEtNTQ2NS00YjQxLTlkM2EtZWVkODYxODYzMmY2%40thread.v2/0?context=%7b%22Tid%22%3a%22909c6581-5130-43e9-88f3-fcb3582cde37%22%2c%22Oid%22%3a%22dc17674c-81d9-4adb-bfb2-8f6a442e4622%22%7d'
```

Retrieve meeting details for given meeting with userName and joinurl

```sh
m365 teams meeting get --userName 'user@tenant.com' --joinUrl 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_MGQ4MDQyNTEtNTQ2NS00YjQxLTlkM2EtZWVkODYxODYzMmY2%40thread.v2/0?context=%7b%22Tid%22%3a%22909c6581-5130-43e9-88f3-fcb3582cde37%22%2c%22Oid%22%3a%22dc17674c-81d9-4adb-bfb2-8f6a442e4622%22%7d'
```

Retrieve meeting details for given meeting with email and joinurl

```sh
m365 teams meeting get --email 'user@tenant.com' --joinUrl 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_MGQ4MDQyNTEtNTQ2NS00YjQxLTlkM2EtZWVkODYxODYzMmY2%40thread.v2/0?context=%7b%22Tid%22%3a%22909c6581-5130-43e9-88f3-fcb3582cde37%22%2c%22Oid%22%3a%22dc17674c-81d9-4adb-bfb2-8f6a442e4622%22%7d'
```

Retrieve meeting details of the currently logged in user with joinurl

```sh
m365 teams meeting get --joinUrl 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_MGQ4MDQyNTEtNTQ2NS00YjQxLTlkM2EtZWVkODYxODYzMmY2%40thread.v2/0?context=%7b%22Tid%22%3a%22909c6581-5130-43e9-88f3-fcb3582cde37%22%2c%22Oid%22%3a%22dc17674c-81d9-4adb-bfb2-8f6a442e4622%22%7d'
```


## Response

=== "JSON"

    ```json
    {
      "id": "AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E_hLMK5kAAAAAAENAABiOC8xvYmdT6G2E_hLMK5kAAIw3TQIAAA=",
      "createdDateTime": "2022-06-26T12:39:34.224055Z",
      "lastModifiedDateTime": "2022-06-26T12:41:36.4357085Z",
      "changeKey": "YjgvMb2JnU+hthPoSzCuZAACMHITIQ==",
      "categories": [],
      "transactionId": null,
      "originalStartTimeZone": "W. Europe Standard Time",
      "originalEndTimeZone": "W. Europe Standard Time",
      "iCalUId": "040000008200E00074C5B7101A82E008000000001AF70ACA5989D801000000000000000010000000048716A892ACAE4DB6CC16097796C401",
      "reminderMinutesBeforeStart": 15,
      "isReminderOn": true,
      "hasAttachments": false,
      "subject": "Test meeting",
      "bodyPreview": "________________________________________________________________________________\r\\\nMicrosoft Teams meeting\r\\\nJoin on your computer or mobile app\r\\\nClick here to join the meeting\r\\\nLearn More | Meeting options\r\\\n___________________________________________",
      "importance": "normal",
      "sensitivity": "normal",
      "isAllDay": false,
      "isCancelled": false,
      "isOrganizer": true,
      "responseRequested": true,
      "seriesMasterId": null,
      "showAs": "busy",
      "type": "singleInstance",
      "webLink": "https://outlook.office365.com/owa/?itemid=AAMkADgzN2Q1NThiLTI0NjYtNGIxYS05MDdjLTg1OWQxNzgwZGM2ZgBGAAAAAAC6jQfUzacTSIHqMw2yacnUBwBiOC8xvYmdT6G2E%2BhLMK5kAAAAAAENAABiOC8xvYmdT6G2E%2BhLMK5kAAIw3TQIAAA%3D&exvsurl=1&path=/calendar/item",
      "onlineMeetingUrl": null,
      "isOnlineMeeting": true,
      "onlineMeetingProvider": "teamsForBusiness",
      "allowNewTimeProposals": true,
      "occurrenceId": null,
      "isDraft": false,
      "hideAttendees": false,
      "responseStatus": {
        "response": "organizer",
        "time": "0001-01-01T00:00:00Z"
      },
      "body": {
        "contentType": "html",
        "content": "<html>\r\\\n<head>\r\\\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\\\n</head>\r\\\n<body>\r\\\n<div><br>\r\\\n<br>\r\\\n<br>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n<div class=\"me-email-text\" lang=\"en-US\" style=\"color:#252424; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n<div style=\"margin-top:24px; margin-bottom:20px\"><span style=\"font-size:24px; color:#252424\">Microsoft Teams meeting</span>\r\\\n</div>\r\\\n<div style=\"margin-bottom:20px\">\r\\\n<div style=\"margin-top:0px; margin-bottom:0px; font-weight:bold\"><span style=\"font-size:14px; color:#252424\">Join on your computer or mobile app</span>\r\\\n</div>\r\\\n<a href=\"https://teams.microsoft.com/l/meetup-join/19%3ameeting_OWIwM2MzNmQtZmY1My00MzM0LWIxMGQtYzkyNzI3OWU5ODMx%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d\" class=\"me-email-headline\" style=\"font-size:14px; font-family:'Segoe UI Semibold','Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif; text-decoration:underline; color:#6264a7\">Click\r\\\n here to join the meeting</a> </div>\r\\\n<div style=\"margin-bottom:24px; margin-top:20px\"><a href=\"https://aka.ms/JoinTeamsMeeting\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">Learn More</a>\r\\\n | <a href=\"https://teams.microsoft.com/meetingOptions/?organizerId=fe36f75e-c103-410b-a18a-2bf6df06ac3a&amp;tenantId=e1dd4023-a656-480a-8a0e-c1b1eec51e1d&amp;threadId=19_meeting_OWIwM2MzNmQtZmY1My00MzM0LWIxMGQtYzkyNzI3OWU5ODMx@thread.v2&amp;messageId=0&amp;language=en-US\" class=\"me-email-link\" style=\"font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\nMeeting options</a> </div>\r\\\n</div>\r\\\n<div style=\"font-size:14px; margin-bottom:4px; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif\">\r\\\n</div>\r\\\n<div style=\"font-size:12px\"></div>\r\\\n</div>\r\\\n<div style=\"width:100%; height:20px\"><span style=\"white-space:nowrap; color:#5F5F5F; opacity:.36\">________________________________________________________________________________</span>\r\\\n</div>\r\\\n<div></div>\r\\\n</body>\r\\\n</html>\r\\\n"
      },
      "start": {
        "dateTime": "2022-06-26T12:30:00.0000000",
        "timeZone": "UTC"
      },
      "end": {
        "dateTime": "2022-06-26T13:00:00.0000000",
        "timeZone": "UTC"
      },
      "location": {
        "displayName": "",
        "locationType": "default",
        "uniqueIdType": "unknown",
        "address": {},
        "coordinates": {}
      },
      "locations": [],
      "recurrence": null,
      "attendees": [
        {
          "type": "required",
          "status": {
            "response": "none",
            "time": "0001-01-01T00:00:00Z"
          },
          "emailAddress": {
            "name": "User D",
            "address": "userD@outlook.com"
          }
        }
      ],
      "organizer": {
        "emailAddress": {
          "name": "User",
          "address": "user@tenant.com"
        }
      },
      "onlineMeeting": {
        "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_OWIwM2MzNmQtZmY1My00MzM0LWIxMGQtYzkyNzI3OWU5ODMx%40thread.v2/0?context=%7b%22Tid%22%3a%22e1dd4023-a656-480a-8a0e-c1b1eec51e1d%22%2c%22Oid%22%3a%22fe36f75e-c103-410b-a18a-2bf6df06ac3a%22%7d"
      }
    }
    ```

=== "Text"

    ```text
    subject : Meeting title
    start   : 26/06/2022, 12:30:00
    end     : 26/06/2022, 13:00:00         
    ```

=== "CSV"

    ```csv
    subject,start,end
    Meeting title,"26/06/2022, 12:30:00","26/06/2022, 13:00:00"
    ```
