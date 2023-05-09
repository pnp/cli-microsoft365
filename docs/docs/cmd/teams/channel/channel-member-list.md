# teams channel member list

Lists members of the specified Microsoft Teams team channel

## Usage

```sh
m365 teams channel member list [options]
```

## Options

`-i, --teamId [teamId]`
: The Id of the Microsoft Teams team. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the Microsoft Teams team. Specify either `teamId` or `teamName` but not both

`-c, --channelId [channelId]`
: The Id of the Microsoft Teams team channel. Specify either `channelId` or `channelName` but not both

`--channelName [channelName]`
: The display name of the Microsoft Teams team channel. Specify either `channelId` or `channelName` but not both

`-r, --role [role]`
: Filter the results to only users with the given role: owner, member, guest

--8<-- "docs/cmd/_global.md"

## Examples
  
List the members of a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000 and channel id 19:00000000000000000000000000000000@thread.skype

```sh
m365 teams channel member list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype
```

List the members of a specified Microsoft Teams team with name _Team Name_ and channel with name _Channel Name_

```sh
m365 teams channel member list --teamName "Team Name" --channelName "Channel Name"
```

List all owners of the specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000 and channel id 19:00000000000000000000000000000000@thread.skype

```sh
m365 teams channel member list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --role owner
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "MCMjMiMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOTpCM25DbkxLd3dDb0dERUFEeVVnUTVrSjVQa2VrdWp5am13eHA3dWhRZUFFMUB0aHJlYWQudGFjdjIjIzc4Y2NmNTMwLWJiZjAtNDdlNC1hYWU2LWRhNWY4YzZmYjE0Mg==",
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
    displayName: John Doe
    email      : johndoe@contoso.onmicrosoft.com
    id         : MCMjMiMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOTpCM25DbkxLd3dDb0dERUFEeVVnUTVrSjVQa2VrdWp5am13eHA3dWhRZUFFMUB0aHJlYWQudGFjdjIjIzc4Y2NmNTMwLWJiZjAtNDdlNC1hYWU2LWRhNWY4YzZmYjE0Mg==
    roles      : ["owner"]
    userId     : 78ccf530-bbf0-47e4-aae6-da5f8c6fb142
    ```

=== "CSV"

    ```csv
    id,roles,displayName,userId,email
    MCMjMiMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOTpCM25DbkxLd3dDb0dERUFEeVVnUTVrSjVQa2VrdWp5am13eHA3dWhRZUFFMUB0aHJlYWQudGFjdjIjIzc4Y2NmNTMwLWJiZjAtNDdlNC1hYWU2LWRhNWY4YzZmYjE0Mg==,"[""owner""]",John Doe,78ccf530-bbf0-47e4-aae6-da5f8c6fb142,johndoe@contoso.onmicrosoft.com
    ```

==="Markdown"

 ```md
# teams channel member list --teamName "Team Name" --channelName "Channel Name"

Date: 5/6/2023

## John Doe (MCMjMiMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOTpCM25DbkxLd3dDb0dERUFEeVVnUTVrSjVQa2VrdWp5am13eHA3dWhRZUFFMUB0aHJlYWQudGFjdjIjIzc4Y2NmNTMwLWJiZjAtNDdlNC1hYWU2LWRhNWY4YzZmYjE0Mg==)

Property | Value
---------|-------
id | MCMjMiMjMGNhYzZjZGEtMmUwNC00YTNkLTljMTYtOWM5MTQ3MGQ3MDIyIyMxOTpCM25DbkxLd3dDb0dERUFEeVVnUTVrSjVQa2VrdWp5am13eHA3dWhRZUFFMUB0aHJlYWQudGFjdjIjIzc4Y2NmNTMwLWJiZjAtNDdlNC1hYWU2LWRhNWY4YzZmYjE0Mg==
roles | ["Owner"]
displayName | John Doe
visibleHistoryStartDateTime | 0001-01-01T00:00:00Z
userId | 78ccf530-bbf0-47e4-aae6-da5f8c6fb142
email | johndoe@contoso.onmicrosoft.com
tenantId | 446355e4-e7e3-43d5-82f8-d7ad8272d55b
```
