# spo group member remove

Removes the specified member from a SharePoint group

## Usage

```sh
m365 spo group member remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the SharePoint group is available

`--groupId [groupId]`
: Id of the SharePoint group from which the user has to be removed. Specify either `groupName` or `groupId`, but not both

`--groupName  [groupName]`
: Name of the SharePoint group from which user has to be removed. Specify either `groupName` or `groupId`, but not both

`--userName <userName>`
: The UPN of the user that needs to be removed (user principal name, eg. megan.bowen@contoso.com)

--8<-- "docs/cmd/_global.md"

## Examples

Remove a user from SharePoint group with id _5_ available on the web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userName "Alex.Wilber@contoso.com"
```

Remove a user from SharePoint group with Name _Site A Visitors_ available on the web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Site A Visitors" --userName "Alex.Wilber@contoso.com"
```
