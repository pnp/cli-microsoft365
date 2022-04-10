# aad groupsetting remove

Removes the particular group setting

## Usage

```sh
m365 aad groupsetting remove [options]
```

## Options

`-i, --id <id>`
: The ID of the group setting to remove

`--confirm`
: Don't prompt for confirming removing the group setting

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified _id_ doesn't refer to an existing group setting, you will get a `Resource does not exist` error.

## Examples

Remove group setting with id _28beab62-7540-4db1-a23f-29a6018a3848_. Will prompt for confirmation before removing the group setting

```sh
m365 aad groupsetting remove --id 28beab62-7540-4db1-a23f-29a6018a3848
```

Remove group setting with id _28beab62-7540-4db1-a23f-29a6018a3848_ without prompting for confirmation

```sh
m365 aad groupsetting remove --id 28beab62-7540-4db1-a23f-29a6018a3848 --confirm
```
