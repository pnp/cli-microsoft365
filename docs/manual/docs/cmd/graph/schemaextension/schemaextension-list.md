# graph schemaextension list

Get a list of schemaExtension objects created in the current tenant, that can be InDevelopment, Available, or Deprecated.

## Usage

```sh
graph schemaextension list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-s, --status [status]`|The status to filter on
`--owner [owner]`|The id of the owner to filter on
`-p, --pageSize [pageSize]`|Number of objects to return
`-n, --pageNumber [pageNumber]`|Page number to return if pageSize is specified (first page is indexed as value of 0)
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get a list of schemaExtension objects created in the current tenant, that can be InDevelopment, Available, or Deprecated.

```sh
graph schemaextension list 
```

Get a list of schemaExtension objects created in the current tenant, with owner 617720dc-85fc-45d7-a187-cee75eaf239e

```sh
graph schemaextension list --owner 617720dc-85fc-45d7-a187-cee75eaf239e
```

## Additional information

pageNumber is specified as a 0-based index. A value of 2 returns the third page of items. 

More information: [https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/schemaextension_list](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/schemaextension_list)
