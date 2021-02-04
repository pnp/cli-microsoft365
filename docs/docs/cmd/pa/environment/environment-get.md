# pa environment get

Gets information about the specified Microsoft Power Apps environment

## Usage

```sh
m365 pa environment get [options]
```

## Options

`-n, --name <name>`
: The name of the environment to get information about

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

## Examples

Get information about the Microsoft Power Apps environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_

```sh
m365 pa environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```
