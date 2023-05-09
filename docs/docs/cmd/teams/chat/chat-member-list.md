# teams chat member list

Lists all members from a Microsoft Teams chat conversation.

## Usage

```sh
m365 teams chat member list [options]
```

## Options

`-i, --chatId <chatId>`
: The ID of the chat conversation

--8<-- "docs/cmd/_global.md"

## Examples

List the members from a Microsoft Teams chat conversation

```sh
m365 teams chat member list --chatId 19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d@unq.gbl.spaces
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "MCMjMCMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOToyNDFhZGJmNi0yYTU2LTRjNzItODFmMi02OWU3NWRlNmFjMzRfNzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyQHVucS5nYmwuc3BhY2VzIyM3OGNjZjUzMC1iYmYwLTQ3ZTQtYWFlNi1kYTVmOGM2ZmIxNDI=",
        "roles": [
          "owner"
        ],
        "displayName": "John Doe",
        "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
        "userId": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
        "email": "johndoe@contoso.onmicrosoft.com",
        "tenantId": "446355e4-e7e3-43d5-82f8-d7ad8272d55b"
      }
    ]
    ```

=== "Text"

    ```text
    userId                                displayName     email
    ------------------------------------  --------------  ---------------------------------
    78ccf530-bbf0-47e4-aae6-da5f8c6fb142  John Doe        johndoe@contoso.onmicrosoft.com
    ```

=== "CSV"

    ```csv
    userId,displayName,email
    78ccf530-bbf0-47e4-aae6-da5f8c6fb142,John Doe,johndoe@contoso.onmicrosoft.com
    ```

=== "Markdown"

    ```md
    # teams chat member list --chatId "19:8b081ef6-4792-4def-b2c9-c363a1bf41d5_5031bb31-22c0-4f6f-9f73-91d34ab2b32d@unq.gbl.spaces"

    Date: 5/8/2023

    ## John Doe (MCMjMCMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOToyNDFhZGJmNi0yYTU2LTRjNzItODFmMi02OWU3NWRlNmFjMzRfNzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyQHVucS5nYmwuc3BhY2VzIyM3OGNjZjUzMC1iYmYwLTQ3ZTQtYWFlNi1kYTVmOGM2ZmIxNDI=)

    Property | Value
    ---------|-------
    id | MCMjMCMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOToyNDFhZGJmNi0yYTU2LTRjNzItODFmMi02OWU3NWRlNmFjMzRfNzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyQHVucS5nYmwuc3BhY2VzIyM3OGNjZjUzMC1iYmYwLTQ3ZTQtYWFlNi1kYTVmOGM2ZmIxNDI=
    roles | ["owner"]
    displayName | John Doe
    visibleHistoryStartDateTime | 0001-01-01T00:00:00Z
    userId | 78ccf530-bbf0-47e4-aae6-da5f8c6fb142
    email | johndoe@contoso.onmicrosoft.com
    tenantId | 446355e4-e7e3-43d5-82f8-d7ad8272d55b
    ```
