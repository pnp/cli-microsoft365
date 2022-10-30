# todo list add

Adds a new Microsoft To Do task list

## Usage

```sh
m365 todo list add [options]
```

## Options

`-n, --name <name>`
: The name of the task list to add.

--8<-- "docs/cmd/_global.md"

## Examples

Add a task list with the name _My task list_

```sh
m365 todo list add --name "My task list"
```

## Response

=== "JSON"

    ```json
    {
      "displayName": "My task list",
      "isOwner": true,
      "isShared": false,
      "wellknownListName": "none",
      "id": "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA="
    }
    ```

=== "Text"

    ```text
    displayName      : My task list
    id               : AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=
    isOwner          : true
    isShared         : false
    wellknownListName: none
    ```
    
=== "CSV"

    ```csv
    displayName,isOwner,isShared,wellknownListName,id
    My task list,1,,none,AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=
    ```

