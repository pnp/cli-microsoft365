# aad tenant id get

Gets Microsoft Azure or Office 365 tenant ID

## Usage

```sh
aad tenant id get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-d, --domainName <domainName>`|The domain name to get the Microsoft Azure or Office 365 tenant ID
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets Microsoft Azure or Office 365 tenant ID

```sh
aad tenant id get --domainName contoso.com
```
