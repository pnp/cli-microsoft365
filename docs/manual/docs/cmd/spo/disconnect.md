# spo disconnect

Disconnects from a previously connected SharePoint Online site

## Usage

```sh
spo disconnect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--verbose`|Runs command with verbose logging

## Remarks

The spo disconnect command disconnects from the previously connected SharePoint Online site and removes any access and refresh tokens from memory.

## Examples

```sh
spo disconnect
```

disconnects from a previously connected SharePoint Online site

```sh
spo disconnect --verbose
```

disconnects from a previously connected SharePoint Online site in verbose mode including detailed debug information in the console output