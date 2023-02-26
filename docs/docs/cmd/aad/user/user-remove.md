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

!!! important
    If the user with the specified id or user name doesn't exist, you will get a `Resource 'xyz' does not exist or one of its queried reference-property objects are not present.` error.

!!! important
    To use this command you must be a Global administrator, User administrator or Privileged Authentication administrator.

!!! note
    After running this command, it may take a minute before the user is effectively moved to the recycle bin.

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
