# todo task list

List tasks from a Microsoft To Do task list

## Usage

```sh
m365 todo task list [options]
```

## Options

`--listName [listName]`
: The name of the task list to return tasks from. Specify either `listName` or `listId`, not both

`--listId [listId]`
: The id of the task list to return tasks from. Specify either `listName` or `listId`, not both

--8<-- "docs/cmd/_global.md"

## Examples

List tasks from Microsoft To Do tasks list with the name _My task list_

```sh
m365 todo task list --listName "My task list"
```

List tasks from Microsoft To Do tasks list with the id AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA==

```sh
m365 todo task list --listId "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=="
```