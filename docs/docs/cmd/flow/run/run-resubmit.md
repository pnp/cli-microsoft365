# flow run resubmit

Resubmits a specific flow run for the specified Microsoft Flow

## Usage

```sh
m365 flow run resubmit [options]
```

## Options

`-n, --name <name>`
: The name of the run to resubmit

`-f, --flow <flow>`
: The name of the flow to resubmit the run for

`-e, --environment <environment>`
: The name of the environment where the Flow is located

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Microsoft Flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error.

If the run with the name you specified doesn't exist, you will get the `The workflow 'xyz' run 'abc' could not be found.` error.

## Examples

Resubmits a specific flow run for the specified Microsoft Flow

```sh
m365 flow run resubmit --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flow 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62
```

Resubmits a specific flow run for the specified Microsoft Flow without prompting for confirmation

```sh
m365 flow run resubmit --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flow 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62 --confirm
```
