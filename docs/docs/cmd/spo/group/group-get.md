# spo group get

Gets site group

## Usage

```sh
m365 spo group get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the group is located

`-i, --id [id]`
: ID of the site group to get. Use either `id` or `name`, but not all. e.g `7`

`--name [name]`
: Name of the site group to get. Specify either `id` or `name` but not both e.g `Team Site Members`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get group with ID _7_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --id 7
```

Get group with name _Team Site Members_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --name "Team Site Members"
```
