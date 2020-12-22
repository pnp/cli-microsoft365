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

--8<-- "docs/cmd/_global.md"

## Examples

Update a task with title _New task_ to _Update doco_ in Microsoft To Do tasks list with the name _My task list_

```sh
m365 todo task set --id "AAMkADU3Y2E0OTMxLTllYTQtNGFlZS1hZGM0LWI1NjZjY2FhM2RhMABGAAAAAADhr7P77n9xS6PdtDemRwpHBwCin1tvQMXzRKN1hQDz2S3VAAAXXsleAACin1tvQMXzRKN1hQDz2S3VAAAXXzr9AAA=" --title "Update doco" --listName "My task list"
```

Update a task with status from _notStarted_ to _inProgress_ in Microsoft To Do tasks list with the name _My task list_

```sh
m365 todo task set --id "AAMkADU3Y2E0OTMxLTllYTQtNGFlZS1hZGM0LWI1NjZjY2FhM2RhMABGAAAAAADhr7P77n9xS6PdtDemRwpHBwCin1tvQMXzRKN1hQDz2S3VAAAXXsleAACin1tvQMXzRKN1hQDz2S3VAAAXXzr9AAA=" --status "inProgress" --listName "My task list"
```