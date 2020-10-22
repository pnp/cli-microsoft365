# aad o365group user add

Adds user to specified Microsoft 365 Group or Microsoft Teams team

## Usage

```sh
m365 aad o365group user add [options]
```

## Alias

```sh
m365 teams user add
```

## Options

`-i, --groupId [groupId]`
: The ID of the Microsoft 365 Group to which to add the user

`--teamId [teamId]`
: The ID of the Teams team to which to add the user

`-n, --userName <userName>`
: User's UPN (user principal name, eg. johndoe@example.com)

`-r, --role [role]`
: The role to be assigned to the new user: `Owner,Member`. Default `Member`

--8<-- "docs/cmd/_global.md"

## Examples

Add a new member to the specified Microsoft 365 Group

```sh
m365 aad o365group user add --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
```

Add a new owner to the specified Microsoft 365 Group

```sh
m365 aad o365group user add --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner
```

Add a new member to the specified Microsoft Teams team

```sh
m365 teams user add --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
```