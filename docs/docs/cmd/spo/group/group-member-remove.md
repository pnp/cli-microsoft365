# spo group member remove

Removes the specified member from a SharePoint group

## Usage

```sh
m365 spo group member remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the SharePoint group is available.

`--groupId [groupId]`
: Id of the SharePoint group from which the user has to be removed. Specify either `groupName` or `groupId`, but not both.

`--groupName  [groupName]`
: Name of the SharePoint group from which user has to be removed. Specify either `groupName` or `groupId`, but not both.

`--userName [userName]`
: The UPN (user principal name, eg. megan.bowen@contoso.com) of the user that needs to be removed. Specify either `userName`, `email`, `userId`, `aadGroupId` or `aadGroupName`.

`--email [email]`
: The email of the user to remove as a member. Specify either `userName`, `email`, `userId`, `aadGroupId` or `aadGroupName`.

`--userId [userId]`
: The user Id (Id of the site user, eg. 14) of the user to remove as a member. Specify either `userName`, `email`, `userId`, `aadGroupId` or `aadGroupName`.
`--aadGroupId [aadGroupId]`
: The object Id of the Azure AD group to remove as a member. Specify either `userName`, `email`, `userId`, `aadGroupId` or `aadGroupName`.

`--aadGroupName [aadGroupName]`
: The name of the Azure AD group to remove as a member. Specify either `userName`, `email`, `userId`, `aadGroupId` or `aadGroupName`.

`--aadGroupId [aadGroupId]`
: The object Id of the Azure AD group to remove as a member. Specify either `aadGroupId`, `aadGroupName` or `userName`

`--aadGroupName [aadGroupName]`
: The name of the Azure AD group to remove as a member. Specify either `aadGroupId`, `aadGroupName` or `userName`

--8<-- "docs/cmd/_global.md"

## Examples

Remove a user from a SharePoint group based on the id on a given web

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userName "Alex.Wilber@contoso.com"
```

Remove a user from a SharePoint group based on the username on a given web

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Site A Visitors" --email "Alex.Wilber@contoso.com"
```

Remove an Azure AD group from a SharePoint group based on the Azure AD group name on a given web

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --aadGroupName "Azure AD Security Group"
```

Remove an Azure AD group from a SharePoint group based on the Azure AD group ID on a given web

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Site A Visitors" --aadGroupId "5786b8e8-c495-4734-b345-756733960730"
```

## Response

The command won't return a response on success.
