# aad o365group conversation post list

Lists conversation posts of a Microsoft 365 group

## Usage

```sh
m365 aad o365group conversation post list [options]
```

## Options

`-i, --groupId [groupId]`
: The Id of the Office 365 Group. You can specify the groupId or groupDisplayName, but not both.

`-d, --groupDisplayName [groupDisplayName]`
: The Displayname of the Office 365 Group. You can specify the groupId or groupDisplayName, but not both.

`-t, --threadId <threadId>`
: The ID of the thread to retrieve details for

--8<-- "docs/cmd/_global.md"

## Examples

Lists the posts of the specific conversation of Microsoft 365 group by groupId

```sh
m365 aad o365group conversation post list --groupId '00000000-0000-0000-0000-000000000000' --threadId 'AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E='
```

Lists the posts of the specific conversation of Microsoft 365 group by groupDisplayName

```sh
m365 aad o365group conversation post list --groupDisplayName 'MyGroup' --threadId 'AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E='
```
