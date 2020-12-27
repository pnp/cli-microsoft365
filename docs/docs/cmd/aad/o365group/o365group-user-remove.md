# aad o365group user remove

Removes the specified user from specified Microsoft 365 Group or Microsoft Teams team

## Usage

```sh
m365 aad o365group user remove [options]
```

## Alias

```sh
m365 aad teams user remove
```

## Options

`-i, --groupId [groupId]`
: The ID of the Microsoft 365 Group from which to remove the user

`--teamId [teamId]`
: The ID of the Microsoft Teams team from which to remove the user

`-n, --userName <userName>`
: User's UPN (user principal name), eg. `johndoe@example.com`

`--confirm`
: Don't prompt for confirming removing the user from the specified Microsoft 365 Group or Microsoft Teams team

--8<-- "docs/cmd/_global.md"

## Remarks

You can remove users from a Microsoft 365 Group or Microsoft Teams team if you are owner of that group or team.

## Examples

Removes user from the specified Microsoft 365 Group

```sh
m365 aad o365group user remove --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
```

Removes user from the specified Microsoft 365 Group without confirmation

```sh
m365 aad o365group user remove --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --confirm
```

Removes user from the specified Microsoft Teams team

```sh
m365 aad teams user remove --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
```