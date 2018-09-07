# spo logout

Log out from SharePoint Online

## Usage

```sh
spo logout [options]
```

## Alias

```sh
spo disconnect
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
    The 'spo disconnect' command is deprecated. Please use 'spo logout' instead.

The `spo logout` command logs out from SharePoint Online and removes any access and refresh tokens from memory.

## Examples

Log out from SharePoint Online

```sh
spo logout
```

Log out from SharePoint Online in debug mode including detailed debug information in the console output

```sh
spo logout --debug
```