# teams app update

Updates Teams app in the organization's app catalog

## Usage

```sh
m365 teams app update [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: ID of the app to update

`-p, --filePath <filePath>`
: Absolute or relative path to the Teams manifest zip file to update in the app catalog

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You can only update a Teams app as a global administrator.

## Examples

Update the Teams app with ID _83cece1e-938d-44a1-8b86-918cf6151957_ from file _teams-manifest.zip_

```sh
m365 teams app update --id 83cece1e-938d-44a1-8b86-918cf6151957 --filePath ./teams-manifest.zip
```