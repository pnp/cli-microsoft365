# graph logout

Log out from the Microsoft Graph

## Usage

```sh
graph logout [options]
```

## Alias

```sh
graph disconnect
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
    The 'graph disconnect' command is deprecated. Please use 'graph logout' instead.

The `graph logout` command logs out from the Microsoft Graph and removes any access and refresh tokens from memory

## Examples

Log out from Microsoft Graph

```sh
graph logout
```

Log out from Microsoft Graph in debug mode including detailed debug information in the console output

```sh
graph logout --debug
```