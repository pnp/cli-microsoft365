# flow export

Exports the specified Microsoft Flow

## Usage

```sh
m365 flow export [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The id of the Microsoft Flow to export

`-e, --environment <environment>`
: The name of the environment for which to export the flow

`-n, --packageDisplayName [packageDisplayName]`
: The display name to use in the exported package

`-d, --packageDescription [packageDescription]`
: The description to use in the exported package

`-c, --packageCreatedBy [packageCreatedBy]`
: The name of the person to be used as the creator of the exported package

`-s, --packageSourceEnvironment [packageSourceEnvironment]`
: The name of the source environment from which the exported package was taken

`-f, --format [format]`
: Export format type. `json,zip`. Default `zip`

`-p, --path [path]`
: The path to save the exported package to

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

If the Microsoft Flow with the id you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error.

## Examples

Export the specified Microsoft Flow as a ZIP file

```sh
m365 flow export --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Export the specified Microsoft Flow as a JSON file

```sh
m365 flow export --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --format json
```

Export the specified Microsoft Flow as a ZIP file, specifying a Display Name of 'My flow name' to be embedded into the package

```sh
m365 flow export --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --packageDisplayName 'My flow name'
```

Export the specified Microsoft Flow as a ZIP file with the filename 'MyFlow.zip' saved to the current directory

```sh
m365 flow export --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --path './MyFlow.zip'
```
