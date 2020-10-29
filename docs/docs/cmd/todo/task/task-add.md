# todo task add

Adds a task to a Microsoft To Do list

## Usage

```sh
m365 todo task add [options]
```

## Options

`-h, --help`
: output usage information

`-t, --title <title>`
: The title of the task

`--listName [listName]`
: The name of the list in which to create the task. Specify either `listName` or `listId` but not both

`--listId [listId]`
: The id of the list in which to create the task. Specify either `listName` or `listId` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Add a task with title _New task_ to Microsoft To Do tasks list with the name _My task list_

```sh
m365 todo task add --title "New task" --listName "My task list"
```

Add a task with title _New task_ to Microsoft To Do tasks list with the id `AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==`

```sh
m365 todo task add --title "New task" --listId "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=="
```
