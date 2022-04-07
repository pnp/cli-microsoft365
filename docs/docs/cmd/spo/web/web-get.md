# spo web get

Retrieve information about the specified site

## Usage

```sh
m365 spo web get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve the information

`--withGroups`
: Set if you want to return associated groups (associatedOwnerGroup, associatedMemberGroup and associatedVisitorGroup) along with other properties

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve information about the site _https://contoso.sharepoint.com/subsite_

```sh
m365 spo web get --webUrl https://contoso.sharepoint.com/subsite
```

Retrieve information about the site _https://contoso.sharepoint.com/subsite_ along with associated groups for the web

```sh
m365 spo web get --webUrl https://contoso.sharepoint.com/subsite --withGroups
```
