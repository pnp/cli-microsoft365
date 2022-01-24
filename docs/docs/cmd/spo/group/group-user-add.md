# spo group user add

Add a user or multiple users to SharePoint Group

## Usage

```sh
m365 spo group user add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the SharePoint group is available

`--groupId [groupId]`
: Id of the SharePoint Group to which user needs to be added, specify either `groupId` or `groupName`

`--groupName [groupName]`
: Name of the SharePoint Group to which user needs to be added, specify either `groupId` or `groupName`

`--userName [userName]`
: User's UPN (user principal name, eg. megan.bowen@contoso.com). If multiple users needs to be added, they have to be comma separated (ex. megan.bowen@contoso.com,alex.wilber@contoso.com), specify either `userName` or `email`

`--email [email]`
: User's email (eg. megan.bowen@contoso.com). If multiple users needs to be added, they have to be comma separated (ex. megan.bowen@contoso.com,alex.wilber@contoso.com), specify either `userName` or `email`

--8<-- "docs/cmd/_global.md"

## Remarks

The command `m365 spo group user add` supports multiple values for the parameter `--userName` or `--email` in a comma seperated way. That being the case, if one of the entries for the parameter `--userName` or `--email` is not valid, the command will fail with error message showing the list of Username/s or email/s that are not valid

## Examples

Add a user with name _Alex.Wilber@contoso.com_ to the SharePoint group with id _5_ available on the web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group user add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userName "Alex.Wilber@contoso.com"
```

Add multiple users by name to the SharePoint group with id _5_ available on the web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group user add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userName "Alex.Wilber@contoso.com, Adele.Vance@contoso.com"
```

Add a user with email _Alex.Wilber@contoso.com_ to the SharePoint group with name _Contoso Site Owners_ available on the web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group user add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Contoso Site Owners" --email "Alex.Wilber@contoso.com"
```

Add multiple users by email to the SharePoint group with name _Contoso Site Owners_ available on the web _https://contoso.sharepoint.com/sites/SiteA_

```sh
m365 spo group user add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Contoso Site Owners" --email "Alex.Wilber@contoso.com, Adele.Vance@contoso.com"
```