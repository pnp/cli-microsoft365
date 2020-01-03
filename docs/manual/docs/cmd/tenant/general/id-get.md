# tenant id get

Gets Office 365 tenant ID for the specified domain

## Usage

```sh
tenant id get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-d, --domainName <domainName>`|The domain name for which to retrieve the Office 365 tenant ID
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get Office 365 tenant ID for the specified domain

```sh
tenant id get --domainName contoso.com
```
