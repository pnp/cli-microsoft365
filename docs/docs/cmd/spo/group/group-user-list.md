# spo group user list

List members of a SharePoint Group

## Usage

```sh
m365 spo group user list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the SharePoint site

`--groupId [groupId]`
: Id of the SharePoint group. Use either `groupName` or `groupId`, but not both

`--groupName [groupName]`
: Name of the SharePoint group. Use either `groupName` or `groupId`, but not both

--8<-- "docs/cmd/_global.md"

## Examples

List the members of the group with ID _5_ for web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group user list --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5
```

List the members of the group with name _Contoso Site Members_ for web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group user list --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Contoso Site Members"
```
