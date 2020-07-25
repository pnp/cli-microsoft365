# teams app publish

Publishes Teams app to the organization's app catalog

## Usage

```sh
m365 teams app publish [options]
```

## Options

`-h, --help`
: output usage information

`-p, --filePath <filePath>`
: Absolute or relative path to the Teams manifest zip file to add to the app catalog

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You can only publish a Teams app as a global administrator.

## Examples

Add the _teams-manifest.zip_ file to the organization's app catalog

```sh
m365 teams app publish --filePath ./teams-manifest.zip
```