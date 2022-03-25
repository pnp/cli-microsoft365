# pa app list

Lists all Power Apps apps

## Usage

```sh
pa app list [options]
```

## Options

`-e, --environment [environment]`
: The name of the environment for which to retrieve available apps

`--asAdmin`
: Set, to list all Power Apps as admin. Otherwise will return only your own apps

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reaches general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

By default, the `app list` command returns only your apps. To list all apps, use the `asAdmin` option and make sure to specify the `environment` option. You cannot specify only one of the options, when specifying the `environment` option the `asAdmin` option has to be present as well.

## Examples

List all your apps

```sh
m365 pa app list
```

List all apps in a given environment

```sh
m365 pa app list --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --asAdmin
```
