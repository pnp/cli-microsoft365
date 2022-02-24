# aad o365group conversation post list

Lists the posts of the specific conversation of Microsoft 365 group

## Usage

```sh
m365 aad o365group conversation post list [options]
```

## Options

`-i, --groupId <groupId>`
: The ID of the Microsoft 365 group

`-t, --threadId <threadId>`
: The ID of the thread to retrieve details for

--8<-- "docs/cmd/_global.md"

## Examples

Lists the posts of the specific conversation of Microsoft 365 group

```sh
m365 aad o365group conversation post list --groupId '00000000-0000-0000-0000-000000000000' --threadId 'AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E='
```
