# teams meeting transcript list

Lists all transcripts for a given meeting.

## Usage

```sh
m365 teams meeting transcript list [options]
```

## Options

`-u, --userId [userId]`
: The id of the user, omit to list attendance reports for current signed in user. Use either  `id`, `userName` or `email`, but not multiple.

`-n, --userName [userName]`
: The name of the user, omit to list attendance reports for current signed in user. Use either `id`, `userName` or `email`, but not multiple.

`--email [email]`
: The email of the user, omit to list attendance reports for current signed in user. Use either `id`, `userName` or `email`, but not multiple.

`-m, --meetingId <meetingId>`
: The Id of the meeting.

--8<-- "docs/cmd/_global.md"

## Examples

Lists all transcripts made for the current signed in user and Microsoft Teams meeting with given id

```sh
m365 teams meeting transcript list --meetingId MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ
```

Lists all transcripts made for the _[garthf@contoso.com](mailto:garthf@contoso.com)_ and Microsoft Teams meeting with given id

```sh
m365 teams meeting transcript list --userName garthf@contoso.com --meetingId MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ
```

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

## Response

=== "JSON"

    ```json
    [
      {
          "id": "MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh",
          "meetingId": "MSpiZTExZjUyMy0yYTRkLTRlYWUtOWQ0Mi0yNzc0MTA4OTNjNDEqMCoqMTk6bWVldGluZ19aakU0WmpVMllqY3RZMkV3T1MwME1UaGtMV0prWlRRdE1qRXhPVGN4T0RaalpUUTJAdGhyZWFkLnYy",
          "meetingOrganizerId": "be11f523-2a4d-4eae-9d42-277410893c41",
          "transcriptContentUrl": "https://graph.microsoft.com/beta/users/be11f523-2a4d-4eae-9d42-277410893c41/onlineMeetings/MSpiZTExZjUyMy0yYTRkLTRlYWUtOWQ0Mi0yNzc0MTA4OTNjNDEqMCoqMTk6bWVldGluZ19aakU0WmpVMllqY3RZMkV3T1MwME1UaGtMV0prWlRRdE1qRXhPVGN4T0RaalpUUTJAdGhyZWFkLnYy/transcripts/MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh/content",
          "createdDateTime": "2021-09-17T06:09:24.8968037Z"
      }
    ]
    ```

=== "Text"

    ```text
    id                                                        createdDateTime
    ------------------------------------                      ---------------------
    MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh  2021-09-17T06:09:24.8968037Z
    ```

=== "CSV"

    ```csv
    id,createdDateTime
    MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh,2021-09-17T06:09:24.8968037Z
    ```

=== "Markdown"

    ```md
    # teams meeting transcript list --meetingId "MSpiZTExZjUyMy0yYTRkLTRlYWUtOWQ0Mi0yNzc0MTA4OTNjNDEqMCoqMTk6bWVldGluZ19aakU0WmpVMllqY3RZMkV3T1MwME1UaGtMV0prWlRRdE1qRXhPVGN4T0RaalpUUTJAdGhyZWFkLnYy"

    Date: 3/25/2023
    
    ## MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh
    
    Property | Value
    ---------|-------
    id | MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh
    meetingId | MSpiZTExZjUyMy0yYTRkLTRlYWUtOWQ0Mi0yNzc0MTA4OTNjNDEqMCoqMTk6bWVldGluZ19aakU0WmpVMllqY3RZMkV3T1MwME1UaGtMV0prWlRRdE1qRXhPVGN4T0RaalpUUTJAdGhyZWFkLnYy
    meetingOrganizerId | be11f523-2a4d-4eae-9d42-277410893c41
    transcriptContentUrl | https://graph.microsoft.com/beta/users/be11f523-2a4d-4eae-9d42-277410893c41/onlineMeetings/MSpiZTExZjUyMy0yYTRkLTRlYWUtOWQ0Mi0yNzc0MTA4OTNjNDEqMCoqMTk6bWVldGluZ19aakU0WmpVMllqY3RZMkV3T1MwME1UaGtMV0prWlRRdE1qRXhPVGN4T0RaalpUUTJAdGhyZWFkLnYy/transcripts/MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh/content
    createdDateTime | 2023-03-25T21:32:08.5586288Z
    ```
