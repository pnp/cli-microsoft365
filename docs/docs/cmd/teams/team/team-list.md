# teams team list

Lists Microsoft Teams teams in the current tenant

## Usage

```sh
m365 teams team list [options]
```

## Options

`-j, --joined`
: Show only joined teams.

--8<-- "docs/cmd/_global.md"

## Remarks

You can only see the details or archived status of the Microsoft Teams you are a member of.

## Examples

List all Microsoft Teams in the tenant.

```sh
m365 teams team list
```

List all Microsoft Teams in the tenant you are a member of.

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
      }
    ]
    ```

=== "Text"

    ``` text
    id                                    displayName       isArchived  description
    ------------------------------------  ----------------  ----------  ---------------------------------------
    5dc7ba76-b9aa-4fdd-9e91-9fe7d0e8dca3  Architecture      false       Architecture Discussion
    ```

=== "CSV"

    ``` text
    id,displayName,isArchived,description
    5dc7ba76-b9aa-4fdd-9e91-9fe7d0e8dca3,Architecture,,Architecture Discussion
    ```

=== "Markdown"

    ```md
    # teams team list --joined "true"

    Date: 1/3/2023

    ## Architecture (5dc7ba76-b9aa-4fdd-9e91-9fe7d0e8dca3)

    Property | Value
    ---------|-------
    id | 5dc7ba76-b9aa-4fdd-9e91-9fe7d0e8dca3
    createdDateTime | null
    displayName | Architecture
    description | Architecture Discussion
    internalId | null
    classification | null
    specialization | null
    visibility | null
    webUrl | null
    isArchived | false
    tenantId | 92e59666-257b-49c3-b1fa-1bae8107f6ba
    isMembershipLimitedToOwners | null
    memberSettings | null
    guestSettings | null
    messagingSettings | null
    funSettings | null
    discoverySettings | null
    summary | null
    ```
