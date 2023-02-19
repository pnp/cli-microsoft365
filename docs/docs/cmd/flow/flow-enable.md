# flow enable

Enables specified Power Automate flow

## Usage

```sh
m365 flow enable [options]
```

## Options

`-n, --name <name>`
: The name of the Power Automate flow to enable

`-e, --environmentName <environmentName>`
: The name of the environment for which to enable Power Automate

`--asAdmin`
: Set, to enable the Power Automate flow as admin

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

By default, the command will try to enable Power Automate flows you own. If you want to enable flows owned by another user, use the `asAdmin` flag.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Power Automate flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error. If you try to enable a non-existing flow as admin, you will get the `Could not find flow 'xyz'.` error.

## Examples

Enables Power Automate flow owned by the currently signed-in user

```sh
m365 flow enable --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Enables Power Automate flow owned by another user

```sh
m365 flow enable --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
```

## Response

The command won't return a response on success.
