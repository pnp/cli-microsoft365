# aad o365group user list

Lists users for the specified Microsoft 365 group

## Usage

```sh
m365 aad o365group user list [options]
```

## Options

`-i, --groupId <groupId>`
: The ID of the Microsoft 365 group for which to list users

`-r, --role [role]`
: Filter the results to only users with the given role: `Owner,Member,Guest`

--8<-- "docs/cmd/_global.md"

## Examples

List all users and their role in the specified Microsoft 365 group

```sh
m365 aad o365group user list --groupId '00000000-0000-0000-0000-000000000000'
```

List all owners and their role in the specified Microsoft 365 group

```sh
m365 aad o365group user list --groupId '00000000-0000-0000-0000-000000000000' --role Owner
```

 List all guests and their role in the specified Microsoft 365 group

```sh
m365 aad o365group user list --groupId '00000000-0000-0000-0000-000000000000' --role Guest
```
