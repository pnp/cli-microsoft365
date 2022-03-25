# teams chat message send

Sends a chat message to a Microsoft Teams chat conversation.

## Usage

```sh
m365 teams chat message send [options]
```

## Options

`--chatId [chatId]`
: The ID of the chat conversation. Specify either `chatId`, `chatName` or `userEmails`, but not multiple.

`--chatName [chatName]`
: The display name of the chat conversation. Specify either `chatId`, `chatName` or `userEmails`, but not multiple.

`-e, --userEmails [userEmails]`
: A comma-separated list of one or more e-mail addresses. Specify either `chatId`, `chatName` or `userEmails`, but not multiple.

`-m, --message <message>`
: The message to send

--8<-- "docs/cmd/_global.md"

## Remarks

A new chat conversation will be created if no existing conversation with the participants specified with emails is found.

## Examples

Send a message to a Microsoft Teams chat conversation by id

```sh
m365 teams chat message send --chatId 19:2da4c29f6d7041eca70b638b43d45437@thread.v2 --message "Welcome to Teams"
```

Send a message to a single person

```sh
m365 teams chat message send --userEmails alexw@contoso.com --message "Welcome to Teams"
```

Send a message to a group of people

```sh
m365 teams chat message send --userEmails alexw@contoso.com,meganb@contoso.com --message "Welcome to Teams"
```

Send a message to a chat conversation finding it by display name

```sh
m365 teams chat message send --chatName "Just a conversation" --message "Welcome to Teams"
```
