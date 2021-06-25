# pa app get

Gets information about the specified Microsoft Power App

## Usage

```sh
pa app get [options]
```

## Options

`-n, --name [name]`
: The name of the Microsoft Power App to get information about

`-d, --displayName [displayName]`
: The display name of the Microsoft Power App to get information about

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reaches general availability.

If you try to retrieve a non-existing Microsoft Power App, you will get the `Request failed with status code 404` error.

## Examples

Get information about the specified Microsoft Power App by the app's name

```sh
m365 pa app get --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Get information about the specified Microsoft Power App by the app's display name

```sh
m365 pa app get --displayName App
```
