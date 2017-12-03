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
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

The spo disconnect command disconnects from the previously connected SharePoint Online site and removes any access and refresh tokens from memory.

## Examples

Disconnect from a previously connected SharePoint Online site

```sh
spo disconnect
```

Disconnects from a previously connected SharePoint Online site in debug mode including detailed debug information in the console output

```sh
spo disconnect --debug
```