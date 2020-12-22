# flow run cancel

Cancels a specific run for the specified flow

## Usage

```sh
m365 flow run cancel [options]
```

## Options

`-n, --name <name>`
: The name of the run to cancel

`-f, --flow <flow>`
: The name of the flow to cancel the run for

`-e, --environment <environment>`
: The name of the environment where the flow is located

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error.

If the run with the name you specified doesn't exist, you will get the `The workflow 'xyz' run 'abc' could not be found.` error.

## Examples

Cancel the given run of the specified flow

```sh
m365 flow run cancel --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flow 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62
```
