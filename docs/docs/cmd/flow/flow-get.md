# flow get

Gets information about the specified Microsoft Flow

## Usage

```sh
m365 flow get [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: The name of the Microsoft Flow to get information about

`-e, --environment <environment>`
: The name of the environment for which to retrieve available Flows

`--asAdmin`
: Set, to retrieve the Flow as admin

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

By default, the command will try to retrieve Microsoft Flows you own. If you want to retrieve Flow owned by another user, use the `asAdmin` flag.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Microsoft Flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error. If you try to retrieve a non-existing flow as admin, you will get the `Could not find flow 'xyz'.` error.

## Examples

Get information about the specified Microsoft Flow owned by the currently signed-in user

```sh
m365 flow get --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Get information about the specified Microsoft Flow owned by another user

```sh
m365 flow get --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
```