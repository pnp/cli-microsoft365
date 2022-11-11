# yammer message remove

Removes a Yammer message

## Usage

```sh
m365 yammer message remove [options]
```

## Options

`--id <id>`
: The id of the Yammer message

`--confirm`
: Don't prompt for confirming removing the Yammer message

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

To remove a message, you must either:

- have posted the message yourself
- be an administrator of the group the message was posted to or
- be an admin of the network the message is in

## Examples

Removes the Yammer message with the id _1239871123_

```sh
m365 yammer message remove --id 1239871123
```

Removes the Yammer message with the id _1239871123_ without prompting for confirmation.

```sh
m365 yammer message remove --id 1239871123 --confirm
```

## Response

The command won't return a response on success.
