# flow disable

Disables specified Power Automate flow

## Usage

```sh
m365 flow disable [options]
```

## Options

`-n, --name <name>`
: The name of the Power Automate flow to disable

`-e, --environment <environment>`
: The name of the environment for which to disable Power Automate flow

`--asAdmin`
: Set, to disable the flow as admin

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

By default, the command will try to disable Power Automate flows you own. If you want to disable Power Automate flows owned by another user, use the `asAdmin` flag.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Power Automate flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error. If you try to disable a non-existing flow as admin, you will get the `Could not find flow 'xyz'.` error.

## Examples

Disables Power Automate flow owned by the currently signed-in user

```sh
m365 flow disable --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Disables Power Automate flow owned by another user

```sh
m365 flow disable --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
```
