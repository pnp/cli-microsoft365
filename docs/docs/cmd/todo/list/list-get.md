# todo list get

Returns a specific Microsoft To Do task list

## Usage

```sh
m365 todo list get [options]
```

## Options

`-i, --id [id]`
: The id of the list. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The name of the list. Specify either `id` or `name` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Get a specific Microsoft To Do task list based on id

```sh
m365 todo list get --id "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA="
```

Get a specific Microsoft To Do task list based on name

```sh
m365 todo list get --name "Task list"
```

## Response

=== "JSON"

    ```json
    {
      "displayName": "Task list",
      "isOwner": true,
      "isShared": false,
      "wellknownListName": "defaultList",
      "id": "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA="
    }
    ```

=== "Text"

    ```text
    displayName    : Task list
    id             : AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=
    ```

=== "CSV"

    ```csv
    displayName,id
    Task list,AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=
    ```
