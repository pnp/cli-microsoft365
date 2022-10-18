# todo task add

Adds a task to a Microsoft To Do list

## Usage

```sh
m365 todo task add [options]
```

## Options

`-t, --title <title>`
: The title of the task.

`--listName [listName]`
: The name of the list in which to create the task. Specify either `listName` or `listId` but not both.

`--listId [listId]`
: The id of the list in which to create the task. Specify either `listName` or `listId` but not both.

`--bodyContent [bodyContent]`
: The body content of the task. In the UI this is called 'notes'.

`--bodyContentType [bodyContentType]`
: The type of the body content. Possible values are `text` and `html`. Default is `text`.

`--dueDateTime [dueDateTime]`
: The date when the task is due. This should be defined as a valid ISO 8601 string in the UTC time zone. Only date value is needed, time value is always ignored.

`--importance [importance]`
: The importance of the task. Possible values are: `low`, `normal`, `high`. Default is `normal`.

`--reminderDateTime [reminderDateTime]`
: The date and time for a reminder alert of the task to occur. This should be defined as a valid ISO 8601 string in the UTC time zone.

--8<-- "docs/cmd/_global.md"

## Examples

Add a task to Microsoft To Do tasks list with with a specific name

```sh
m365 todo task add --title "New task" --listName "My task list"
```

Add a task to a Microsoft To Do tasks list with a specific id

```sh
m365 todo task add --title "New task" --listId "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=="
```

Create a new task with bodyContent and reminder and flag it as important

```sh
m365 todo task add --title "New task" --listName "My task list" --bodyContent "I should not forget this" --reminderDateTime 2023-01-01T12:00:00Z --importance high
```

Create a new task with a specific due date

```sh
m365 todo task add --title "New task" --listId "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==" --dueDateTime 2023-01-01
```
