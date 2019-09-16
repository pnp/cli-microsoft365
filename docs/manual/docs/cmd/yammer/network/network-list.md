# Yammer network list

Returns a list of networks to which the current user has access

## Usage

```sh
yammer network list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--includeSuspended true|false`|Include the networks the user is suspended. Default `false`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Returns a list of networks to which the current user has access.

```sh
yammer network list
```

Returns a list of networks to which the current user has access including the networks the user is suspended.

```sh
yammer network list --includeSuspended true
```