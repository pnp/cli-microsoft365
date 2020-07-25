# flow run get

Gets information about a specific run of the specified Microsoft Flow

## Usage

```sh
m365 flow run get [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: The name of the run to get information about

`-f, --flow <flow>`
: The name of the Microsoft Flow for which to retrieve information

`-e, --environment <environment>`
: The name of the environment where the Flow is located

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

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Microsoft Flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error.

If the run with the name you specified doesn't exist, you will get the `The provided workflow run name is not valid.` error.

## Examples

Get information about the given run of the specified Microsoft Flow

```sh
m365 flow run get --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --flow 5923cb07-ce1a-4a5c-ab81-257ce820109a --name 08586653536760200319026785874CU62
```