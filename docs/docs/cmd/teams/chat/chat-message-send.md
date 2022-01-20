# teams chat message send

Sends a chat message to a Microsoft Teams chat conversation.

## Usage

```sh
m365 teams chat message send [options]
```

## Options

`--chatId [chatId]`
: The ID of the chat conversation. Specify either `chatId`, `chatName` or `userEmails`, but not two or all three.

`--chatName [chatName]`
: The display name of the chat conversation. Specify either `chatId`, `chatName` or `userEmails`, but not two or all three. 

`-e, --userEmails [userEmails]`
: A comma-separated list of one or more e-mailaddresses. Specify either `chatId`, `chatName` or `userEmails`, but not two or all three. A new chat conversation will be created if no existing conversation with these participants can be found.

`-m, --message <message>`
: The message to send

--8<-- "docs/cmd/_global.md"

## Examples

Send a message to a Microsoft Teams chat conversation by Id

```sh
m365 teams chat message send --chatId 19:2da4c29f6d7041eca70b638b43d45437@thread.v2
```

Send a message to a single person

```sh
m365 teams chat message send --userEmails alexw@contoso.com
```

Send a message to a group of people

```sh
m365 teams chat message send --userEmails alexw@contoso.com,meganb@contoso.com
```

Send a message to a chat conversation finding it by display name

```sh
m365 teams chat message send --chatName "Just a conversation"
```