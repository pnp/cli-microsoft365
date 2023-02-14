# aad user remove

Removes a specific user

## Usage

```sh
m365 aad user remove [options]
```

## Options

`--id [id]`
: The ID of the user. Specify either `id` or `userName` but not both.

`--userName [userName]`
:	User principal name of the user. Specify either `id` or `userName` but not both.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Remarks

If the user with the specified id or user name doesn't exist, you will get a `Resource 'xyz' does not exist or one of its queried reference-property objects are not present.` error.

For removing a user you need one of the following roles:
- User Administrator
- Privileged Authentication Administrator
- Global Administrator

## Examples

Removes a specific user by id

```sh
m365 aad user remove --id a33bd401-9117-4e0e-bb7b-3f61c1539e10
```

Removes a specific user by its UPN

```sh
m365 aad user remove --name john.doe@contoso.com
```

## Response

The command won't return a response on success.
