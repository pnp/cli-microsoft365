# yammer message add

Posts a Yammer network message on behalf of the current user

## Usage

```sh
m365 yammer message add [options]
```

## Options

`-b, --body <body>`
: The text of the message body

`--groupId [groupId]`
: Post the message to this group, specified by ID. If this is set then the networkId is inferred from it. You must either specify `groupId`, `repliedToId`, or `directToUserIds` to send the message

`-r, --repliedToId [repliedToId]`
: The message ID this message is in reply to. If this is set then groupId and networkId are inferred from it. You must either specify `groupId`, `repliedToId`, or `directToUserIds` to send the message

`-d, --directToUserIds [directToUserIds]`
: Send a private message to one or more users, specified by ID. Alternatively, you can use the Yammer network e-mail addresses instead of the IDs. You must either specify `groupId`, `repliedToId`, or `directToUserIds` to send the message

`--networkId [networkId]`
: Specify the network to post a message

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

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

Posts a message to the group with the ID 12312312312 in the network 11112312

```sh
m365 yammer message add --body "Hello everyone!" --groupId 12312312312 --networkId 11112312
```
