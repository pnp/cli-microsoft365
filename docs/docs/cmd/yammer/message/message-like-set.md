# yammer message like set

Likes or unlikes a Yammer message

## Usage

```sh
m365 yammer message like set [options]
```

## Options

`-h, --help`
: output usage information

`--id <id>`
: The id of the Yammer message

`--enable [enable]`
: Set to `true` to like a message. Set to `false` to unlike it. Default `true`

`--confirm`
: Don't prompt for confirmation before unliking a message

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Likes the message with the ID `5611239081`

```sh
m365 yammer message like set --id 5611239081
```

Unlike the message with the ID `5611239081`

```sh
m365 yammer message like set --id 5611239081 --enable false
```

Unlike the message with the ID `5611239081` without asking for confirmation

```sh
m365 yammer message like set --id 5611239081 --enable false --confirm
```
