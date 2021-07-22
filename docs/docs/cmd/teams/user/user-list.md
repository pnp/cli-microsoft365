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
