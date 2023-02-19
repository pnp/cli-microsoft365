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

!!! important
    To use this command you must be a Global administrator, User administrator or Privileged Authentication administrator

!!! note
    After running this command, it may take a minute before all deleted users are effectively removed from the tenant.

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
