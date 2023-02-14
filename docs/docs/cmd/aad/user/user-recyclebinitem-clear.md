# aad user recyclebinitem clear

Removes all users from the tenant recycle bin

## Usage

```sh
m365 aad user recyclebinitem clear [options]
```

## Options

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Remarks

You will need one of the following roles assigned to permanently delete a user: User Administrator, Privileged Authentication Administrator or Global administrator.

## Examples

Removes all users from the tenant recycle bin

```sh
m365 aad user recyclebinitem clear
```

Removes all users from the tenant recycle bin without confirmation prompt

```sh
m365 aad user recyclebinitem clear --confirm
```

## Response

The command won't return a response on success.
