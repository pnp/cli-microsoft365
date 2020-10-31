# aad o365group user set

Updates role of the specified user in the specified Microsoft 365 Group or Microsoft Teams team

## Usage

```sh
m365 aad o365group user set [options]
```

## Alias

```sh
m365 aad teams user set
```

## Options

`-h, --help`
: output usage information

`-i, --groupId [groupId]`
: The ID of the Microsoft 365 group for which to update user

`--teamId [teamId]`
: The ID of the Microsoft Teams team for which to update user

`-n, --userName <userName>`
: UPN of the user for whom to update the role (eg. johndoe@example.com)

`-r, --role <role>`
: Role to set for the given user in the specified Microsoft 365 Group or Microsoft Teams team. Allowed values: `Owner,Member`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

The command will return an error if the user already has the specified role in the given Microsoft 365 Group or Microsoft Teams team.

## Examples

Promote the specified user to owner of the given Microsoft 365 Group

```sh
m365 aad o365group user list --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner
```

Demote the specified user from owner to member in the given Microsoft 365 Group

```sh
m365 aad o365group user list --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Member
```

Promote the specified user to owner of the given Microsoft Teams team

```sh
m365 aad teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner
```

Demote the specified user from owner to member in the given Microsoft Teams team

```sh
m365 aad teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Member
```
