# spo contenttypehub get

Returns the URL of the SharePoint Content Type Hub of the Tenant

## Usage

```sh
spo contenttypehub get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples
  
Retrieve the Content Type Hub URL

```sh
spo contenttypehub get
```
