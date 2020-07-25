# yammer message add

Posts a Yammer network message on behalf of the current user

## Usage

```sh
m365 yammer message add [options]
```

## Options

`-h, --help`
: output usage information

`-b, --body <body>`
: The text of the message body

`-r, --repliedToId [repliedToId]`
: The message ID this message is in reply to. If this is set then groupId and networkId are inferred from it

`-d, --directToUserIds [directToUserIds]`
: Send a private message to one or more users, specified by ID. Alternatively, you can use the Yammer network e-mail addresses instead of the IDs

`--groupId [groupId]`
: Post the message to this group, specified by ID. If this is set then the networkId is inferred from it. A post without directToUserIds, repliedToId or groupId will default to All Company group

`--networkId [networkId]`
: Post a message in the "All Company" feed of this network, if repliedToId, directToUserIds and groupId are all omitted

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

Posts a message to the "All Company" feed

```sh
m365 yammer message add --body "Hello everyone!"
```

Replies to a message with the ID 1231231231

```sh
m365 yammer message add --body "Hello everyone!" --repliedToId 1231231231
```

Sends a private conversation to the user with the ID 1231231231

```sh
m365 yammer message add --body "Hello everyone!" --directToUserIds 1231231231
```

Sends a private conversation to multiple users by ID

```sh
m365 yammer message add --body "Hello everyone!" --directToUserIds "1231231231,1121312"
```

Sends a private conversation to the user with the e-mail pl@nubo.eu and sc@nubo.eu

```sh
m365 yammer message add --body "Hello everyone!" --directToUserIds "pl@nubo.eu,sc@nubo.eu"
```

Posts a message to the group with the ID 12312312312

```sh
m365 yammer message add --body "Hello everyone!" --groupId 12312312312
```

Posts a message to the "All Company" feed of the network 11112312

```sh
m365 yammer message add --body "Hello everyone!" --networkId 11112312
```
