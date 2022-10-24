# todo task set

Update a task in a Microsoft To Do task list

## Usage

```sh
m365 todo task set [options]
```

## Options

`-i, --id <id>`
: The id of the task to update

`-t, --title [title]`
: Sets the task title

`-s, --status [status]`
: Sets the status of the task. Allowed values `notStarted`,`inProgress`,`completed`,`waitingOnOthers`,`deferred`

`--listName [listName]`
: The name of the task list in which the task exists. Specify either `listName` or `listId`, not both

`--listId [listId]`
: The id of the task list in which the task exists. Specify either `listName` or `listId`, not both

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
