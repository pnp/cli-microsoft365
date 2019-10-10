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
`--includeSuspended`|Include the networks in which the user is suspended
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Returns the current user's networks

```sh
yammer network list
```

Returns the current user's networks including the networks in which the user is suspended

```sh
yammer network list --includeSuspended
```