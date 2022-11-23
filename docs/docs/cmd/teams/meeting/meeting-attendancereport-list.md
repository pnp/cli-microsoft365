# teams meeting attendancereport list

Lists all attendance reports for a given meeting

## Usage

```sh
m365 teams meeting attendancereport list [options]
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

Lists all attendance reports made for the current signed in user and Microsoft Teams meeting with given id

```sh
m365 teams meeting attendancereport list --meetingId _MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ_
```

Lists all attendance reports made for the garthf@contoso.com and Microsoft Teams meeting with given id

```sh
m365 teams meeting attendancereport list --userName garthf@contoso.com --meetingId MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "ae6ddf54-5d48-4448-a7a9-780eee17fa13",
        "totalParticipantCount": 6,
        "meetingStartDateTime": "2022-11-22T22:46:46.981Z",
        "meetingEndDateTime": "2022-11-22T22:47:07.703Z"
      }
    ]
    ```

=== "Text"

    ```text
    id                                    totalParticipantCount
    ------------------------------------  ---------------------
    ae6ddf54-5d48-4448-a7a9-780eee17fa13  6
    ```

=== "CSV"

    ```csv
    id,totalParticipantCount
    ae6ddf54-5d48-4448-a7a9-780eee17fa13,6
    ```
