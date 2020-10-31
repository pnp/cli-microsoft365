# pa connector export

Exports the specified power automate or power apps custom connector

## Usage

```sh
m365 pa connector export [options]
```

## Alias

```sh
m365 flow connector export
```

## Options

`-h, --help`
: output usage information

`-e, --environment <environment>`
: The name of the environment where the custom connector to export is located

`-c, --connector <connector>`
: The name of the custom connector to export

`--outputFolder [outputFolder]`
: Path where the folder with connector's files should be saved. If not specified, will create the connector's folder in the current folder.

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If no output folder has been specified, the `pa connector export` command will create a folder named after the connector in the current folder. If the output folder has been specified, the folder named after the connector will be created in that folder.

## Examples

Export the specified custom connector

```sh
m365 pa connector export --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --connector shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa
```

Export the specified custom connector to the specific directory

```sh
m365 pa connector export --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --connector shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa --outputFolder connector
```
