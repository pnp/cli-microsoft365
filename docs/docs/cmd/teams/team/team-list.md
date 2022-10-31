# teams team list

Lists Microsoft Teams teams in the current tenant

## Usage

```sh
m365 teams team list [options]
```

## Options

`-j, --joined`
: Show only joined teams

--8<-- "docs/cmd/_global.md"

## Remarks

You can only see the details or archived status of the Microsoft Teams you are a member of.

## Examples

List all Microsoft Teams in the tenant

```sh
m365 teams team list
```

List all Microsoft Teams in the tenant you are a member of

```sh
m365 teams team list --joined
```

## Response

=== "JSON"

    ``` json
    [
      {
        "id": "5dc7ba76-b9aa-4fdd-9e91-9fe7d0e8dca3",
        "displayName": "Architecture",
        "isArchived": false,
        "description": "Architecture Discussion"
      },
      {
        "id": "9234218f-3682-4ccc-88b2-af223443beb1",
        "displayName": "Architecture2",
        "isArchived": false,
        "description": "Architecture Discussion2"
      }
    ]
    ```

=== "Text"

    ``` text
    id                                    displayName       isArchived  description
    ------------------------------------  ----------------  ----------  ---------------------------------------
    5dc7ba76-b9aa-4fdd-9e91-9fe7d0e8dca3  Architecture      false       Architecture Discussion
    9234218f-3682-4ccc-88b2-af223443beb1  Architecture2     false       Architecture Discussion2
    ```

=== "CSV"

    ``` text
    id,displayName,isArchived,description
    5dc7ba76-b9aa-4fdd-9e91-9fe7d0e8dca3,Architecture,,Architecture Discussion
    9234218f-3682-4ccc-88b2-af223443beb1,Architecture2,,Architecture Discussion2
    ```
