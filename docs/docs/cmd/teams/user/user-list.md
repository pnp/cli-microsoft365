# teams user list

Lists users for the specified Microsoft Teams team

## Usage

```sh
m365 teams user list [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the Microsoft Teams team for which to list users

`-r, --role [role]`
: Filter the results to only users with the given role: `Owner,Member,Guest`

--8<-- "docs/cmd/_global.md"

## Examples

List all users and their role in the specified Microsoft teams team

```sh
m365 teams user list --teamId '00000000-0000-0000-0000-000000000000'
```

List all owners and their role in the specified Microsoft teams team

```sh
m365 teams user list --teamId '00000000-0000-0000-0000-000000000000' --role Owner
```

## Response

=== "JSON"

    ``` json
    [
      {
        "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
        "displayName": "Nico De Cleyre",
        "userPrincipalName": "nico@contoso.onmicrosoft.com",
        "userType": "Owner"
      }
    ]
    ```

=== "Text"

    ``` text
    displayName      : Nico De Cleyre
    id               : 78ccf530-bbf0-47e4-aae6-da5f8c6fb142
    userPrincipalName: nico@contoso.onmicrosoft.com
    userType         : Owner
    ```

=== "CSV"

    ``` text
    id,displayName,userPrincipalName,userType
    78ccf530-bbf0-47e4-aae6-da5f8c6fb142,Nico De Cleyre,nico@contos.onmicrosoft.com,Owner
    ```
