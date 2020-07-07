# yammer message remove

Removes a Yammer message

## Usage

```sh
yammer message remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--id <id>`|The id of the Yammer message
`--confirm`|Don't prompt for confirming removing the Yammer message
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the Office 365 CLI the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

To remove a message, you must either:

- have posted the message yourself
- be an administrator of the group the message was posted to or
- be an admin of the network the message is in

## Examples

Removes the Yammer message with the id _1239871123_

```sh
yammer message remove --id 1239871123
```

Removes the Yammer message with the id _1239871123_ without prompting for confirmation.

```sh
yammer message remove --id 1239871123 --confirm
```
