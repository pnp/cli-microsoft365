# app remove

Removes the specified Power App

## Usage

```sh
m365 pa app remove [options]
```

## Options

`-n, --name <name>`
: The name of the Power App to remove

`-e, --environment <environment>`
: The name of the environment to which the Power App belongs

`--asAdmin`
: Set, to remove the Power App as admin

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Remarks

By default, the command will try to remove a Power App you own. If you want to remove a Power App owned by another user, use the `asAdmin` flag.

If you want to remove a model-driven Power App you will always need the `asAdmin` flag.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Power App with the name you specified doesn't exist, you will get the `Error: Resource 'abc' does not exist in environment 'xyz'` error.

## Examples

Removes the specified Power App owned by the currently signed-in user

```sh
m365 pa app remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Removes the specified Power App owned by the currently signed-in user without confirmation

```sh
m365 pa app remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --confirm
```

Removes the specified Power App owned by another user

```sh
m365 pa app remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
```

Removes the specified Power App owned by another user without confirmation

```sh
m365 pa app remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin --confirm
```
