# todo task set

Update a task in a Microsoft To Do task list

## Usage

```sh
m365 todo task set [options]
```

## Options

`-i, --id <id>`
: The id of the task to update.

`-t, --title [title]`
: Sets the task title.

`-s, --status [status]`
: Sets the status of the task. Allowed values are `notStarted`, `inProgress`, `completed`, `waitingOnOthers`, `deferred`.

`--listName [listName]`
: The name of the task list in which the task exists. Specify either `listName` or `listId`, but not both.

`--listId [listId]`
: The id of the task list in which the task exists. Specify either `listName` or `listId`, but not both.

`--bodyContent [bodyContent]`
: The body content of the task. In the UI this is called 'notes'.

`--bodyContentType [bodyContentType]`
: The type of the body content. Possible values are `text` and `html`. Default is `text`.

`--importance [importance]`
: The importance of the task. Possible values are: `low`, `normal`, `high`.

`--dueDateTime [dueDateTime]`
: The date and time when the task is due. This should be defined as a valid ISO 8601 string in the UTC time zone. Only date value is needed, time value is always ignored.

`--reminderDateTime [reminderDateTime]`
: The date and time for a reminder alert of the task to occur. This should be defined as a valid ISO 8601 string in the UTC time zone.

--8<-- "docs/cmd/_global.md"

## Examples

Update a task with title _New task_ to _Update doco_ in Microsoft To Do tasks list with a specific name

```sh
m365 todo task set --id "AAMkADU3Y2E0OTMxLTllYTQtNGFlZS1hZGM0LWI1NjZjY2FhM2RhMABGAAAAAADhr7P77n9xS6PdtDemRwpHBwCin1tvQMXzRKN1hQDz2S3VAAAXXsleAACin1tvQMXzRKN1hQDz2S3VAAAXXzr9AAA=" --title "Update doco" --listName "My task list"
```

Update a task with status from _notStarted_ to _inProgress_ in Microsoft To Do tasks list with a specific name

```sh
m365 todo task set --id "AAMkADU3Y2E0OTMxLTllYTQtNGFlZS1hZGM0LWI1NjZjY2FhM2RhMABGAAAAAADhr7P77n9xS6PdtDemRwpHBwCin1tvQMXzRKN1hQDz2S3VAAAXXsleAACin1tvQMXzRKN1hQDz2S3VAAAXXzr9AAA=" --status "inProgress" --listName "My task list"
```

Update a task with bodyContent and reminder and flag it as important in Microsoft To Do tasks list with a specific name

```sh
m365 todo task set --id "AAMkADU3Y2E0OTMxLTllYTQtNGFlZS1hZGM0LWI1NjZjY2FhM2RhMABGAAAAAADhr7P77n9xS6PdtDemRwpHBwCin1tvQMXzRKN1hQDz2S3VAAAXXsleAACin1tvQMXzRKN1hQDz2S3VAAAXXzr9AAA=" --listName "My task list" --bodyContent "I should not forget this" --reminderDateTime 2023-01-01T12:00:00Z --importance high
```

Update a task with due date in Microsoft To Do tasks list with list id

```sh
m365 todo task set --id "AAMkADU3Y2E0OTMxLTllYTQtNGFlZS1hZGM0LWI1NjZjY2FhM2RhMABGAAAAAADhr7P77n9xS6PdtDemRwpHBwCin1tvQMXzRKN1hQDz2S3VAAAXXsleAACin1tvQMXzRKN1hQDz2S3VAAAXXzr9AAA=" --listId "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==" --dueDateTime 2023-01-01
```

## Response

=== "JSON"

    ```json
    {
      "importance": "high",
      "isReminderOn": true,
      "status": "notStarted",
      "title": "Update doco",
      "createdDateTime": "2022-10-29T11:03:20.9175176Z",
      "lastModifiedDateTime": "2022-10-30T14:07:03.0718199Z",
      "hasAttachments": false,
      "categories": [],
      "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtSAAA=",
      "body": {
        "content": "I should not forget this",
        "contentType": "text"
      },
      "dueDateTime": {
        "dateTime": "2023-01-01T00:00:00.0000000",
        "timeZone": "UTC"
      },
      "reminderDateTime": {
        "dateTime": "2023-01-01T12:00:00.0000000",
        "timeZone": "UTC"
      }
    }
	  ```

=== "Text"

    ```text
    body                : {"content":"I should not forget this","contentType":"text"}
    categories          : []
    createdDateTime     : 2022-10-29T11:03:20.9175176Z
    dueDateTime         : {"dateTime":"2023-01-01T00:00:00.0000000","timeZone":"UTC"}
    hasAttachments      : false
    id                  : AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtSAAA=
    importance          : high
    isReminderOn        : true
    lastModifiedDateTime: 2022-10-30T14:08:17.6665299Z
    reminderDateTime    : {"dateTime":"2023-01-01T12:00:00.0000000","timeZone":"UTC"}
    status              : notStarted
    title               : Update doco
	  ```

=== "CSV"

    ```csv
    importance,isReminderOn,status,title,createdDateTime,lastModifiedDateTime,hasAttachments,categories,id,body,dueDateTime,reminderDateTime
    high,1,notStarted,Update doco,2022-10-29T11:03:20.9175176Z,2022-10-30T14:09:14.7687057Z,,[],AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtSAAA=,"{""content"":""I should not forget this"",""contentType"":""text""}","{""dateTime"":""2023-01-01T00:00:00.0000000"",""timeZone"":""UTC""}","{""dateTime"":""2023-01-01T12:00:00.0000000"",""timeZone"":""UTC""}"
	  ```

