# aad disconnect

Disconnects from Azure Active Directory Graph

## Usage

```sh
aad disconnect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

The `aad disconnect` command disconnects from Azure Active Directory Graph and removes any access and refresh tokens from memory.

## Examples

Disconnect from Azure Active Directory Graph

```sh
aad disconnect
```

Disconnect from Azure Active Directory Graph in debug mode including detailed debug information in the console output

```sh
aad disconnect --debug
```