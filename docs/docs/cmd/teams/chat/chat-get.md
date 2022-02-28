# teams chat get

Get a Microsoft Teams chat conversation.

## Usage

```sh
m365 teams chat get [options]
```

## Options

`-i, --id [id]`
: The ID of the chat conversation. Specify either `id`, `name` or `participants`, but not multiple.

`-n, --name [name]`
: The display name of the chat conversation. Specify either `id`, `name` or `participants`, but not multiple.

`-p, --participants [participants]`
: A comma-separated list of one or more e-mail addresses. Specify either `id`, `name` or `participants`, but not multiple.

--8<-- "docs/cmd/_global.md"

## Remarks

The output will not include the chat conversation members or messages. It will just retrieve the conversation details.

## Examples

Get a Microsoft Teams chat conversation by id

```sh
m365 teams chat get --id 19:2da4c29f6d7041eca70b638b43d45437@thread.v2
```

Get a Microsoft Teams one on one chat conversation, finding it by participant.

```sh
m365 teams chat get --participants alexw@contoso.com
```

Get a Microsoft Teams group chat conversation, finding it by participants.

```sh
m365 teams chat get --participants alexw@contoso.com,meganb@contoso.com
```

Get a Microsoft Teams chat conversation, finding it by display name

```sh
m365 teams chat get --name "Just a conversation"
```
