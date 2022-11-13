# yammer message like set

Likes or unlikes a Yammer message

## Usage

```sh
m365 yammer message like set [options]
```

## Options

`--messageId <messageId>`
: The id of the Yammer message

`--enable [enable]`
: Set to `true` to like a message. Set to `false` to unlike it. Default `true`

`--confirm`
: Don't prompt for confirmation before unliking a message

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Likes the message with the ID `5611239081`

```sh
m365 yammer message like set --messageId 5611239081
```

Unlike the message with the ID `5611239081`

```sh
m365 yammer message like set --messageId 5611239081 --enable false
```

Unlike the message with the ID `5611239081` without asking for confirmation

```sh
m365 yammer message like set --messageId 5611239081 --enable false --confirm
```

## Response

The command won't return a response on success.
