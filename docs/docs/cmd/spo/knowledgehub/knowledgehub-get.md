# spo knowledgehub get

Gets the Knowledge Hub Site URL for your tenant

## Usage

```sh
m365 spo knowledgehub get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Gets the Knowledge Hub Site URL for your tenant

```sh
m365 spo knowledgehub get
```
