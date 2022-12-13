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
: The UPN (user principal name, eg. megan.bowen@contoso.com) of the user that needs to be removed. Specify either `userName`, `email`, or `userId`, but not multiple.

`--email [email]`
: The email of the user to remove as a member. Specify either `userName`, `email`, or `userId`, but not multiple.

`--userId [userId]`
: The user Id (Id of the site user, eg. 14) of the user to remove as a member. Specify either `userName`, `email`, or `userId`, but not multiple.

--8<-- "docs/cmd/_global.md"

## Examples

Remove a user by id from SharePoint group on the web

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userName "Alex.Wilber@contoso.com"
```

Remove a user by email from SharePoint group on the web

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Site A Visitors" --email "Alex.Wilber@contoso.com"
```

Remove a user by id from SharePoint group on the web

```sh
m365 spo group member remove --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Site A Visitors" --userId 14
```

## Response

The command won't return a response on success.
