# aad logout

Log out from Azure Active Directory Graph

## Usage

```sh
aad logout [options]
```

## Alias

```sh
aad disconnect
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    The 'aad disconnect' command is deprecated. Please use 'aad logout' instead.

The `aad logout` command logs out from Azure Active Directory Graph and removes any access and refresh tokens from memory.

## Examples

Log out from Azure Active Directory Graph

```sh
aad logout
```

Log out from Azure Active Directory Graph in debug mode including detailed debug information in the console output

```sh
aad logout --debug
```