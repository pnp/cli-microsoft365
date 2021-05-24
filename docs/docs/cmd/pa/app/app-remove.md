# pa app remove

Removes the specified Power App

## Usage

```sh
m365 pa app remove [options]
```

## Options

`-n, --name <name>`
: The name of the Power App to remove

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Remarks

By default, the command will try to remove a Power App. As maker, you are able to delete the Power Apps you own. As administrator, you are also able to delete Power Apps from other users.

To remove a model-driven Power App you need administrator permissions.

If the Power App with the name you specified doesn't exist, you will get the `Error: App 'abc' does not exist` error.

## Examples

Removes the specified Power App

```sh
m365 pa app remove --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Removes the specified Power App without confirmation

```sh
m365 pa app remove --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --confirm
```
