# spo web get

Retrieve information about the specified site

## Usage

```sh
spo web get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site for which to retrieve the information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Retrieve information about the site _https://contoso.sharepoint.com/subsite_

```sh
spo web get --webUrl https://contoso.sharepoint.com/subsite
```