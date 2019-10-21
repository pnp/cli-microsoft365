# yammer message list

Returns all accessible messages from the user's Yammer network

## Usage

```sh
yammer message list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--olderThanId [olderThanId]`|Returns messages older than the message ID specified as a numeric string
`--threaded`|Will only return the thread starter (first message) for each thread. This parameter is intended for apps which need to display message threads collapsed
`--limit [limit]`|Limits the messages returned
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the Office 365 CLI the permission to the Yammer API. To do this, execute the `consent --service yammer` command.

## Examples

Returns all Yammer network messages

```sh
yammer message list
```

Returns all Yammer network messages older than the message ID 5611239081

```sh
yammer message list --olderThanId 5611239081
```

Returns all Yammer network thread starter (first message) for each thread

```sh
yammer message list --threaded
```

Returns the first 10 Yammer network messages

```sh
yammer message list --limit 10
```
