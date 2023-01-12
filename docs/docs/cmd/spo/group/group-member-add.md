# spo group member add

Add members to a SharePoint Group

## Usage

```sh
m365 spo group member add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the SharePoint group is available

`--groupId [groupId]`
: Id of the SharePoint Group to which the user needs to be added, specify either `groupId` or `groupName`

`--groupName [groupName]`
: Name of the SharePoint Group to which the user needs to be added, specify either `groupId` or `groupName`

`--userName [userName]`
: User's UPN (user principal name, eg. megan.bowen@contoso.com). If multiple users need to be added, they have to be comma separated (ex. megan.bowen@contoso.com,alex.wilber@contoso.com). Specify either `userId`, `userName`, `email`, `aadGroupId` or `aadGroupName`

`--email [email]`
: User's email (eg. megan.bowen@contoso.com). If multiple users need to be added, they have to be comma separated (ex. megan.bowen@contoso.com,alex.wilber@contoso.com). Specify either `userId`, `userName`, `email`, `aadGroupId` or `aadGroupName`

`--userId [userId]`
: The user Id of the user to add as a member. (Id of the site user, for example: 14) If multiple users need to be added, the Ids have to be comma separated. Specify either `userId`, `userName`, `email`, `aadGroupId` or `aadGroupName`

`--aadGroupId [aadGroupId]`
: The object Id of the Azure AD group to add as a member. If multiple groups need to be added, the Ids have to be comma separated. Specify either `userId`, `userName`, `email`, `aadGroupId` or `aadGroupName`

`--aadGroupName [aadGroupName]`
: The name of the Azure AD group to add as a member. Specify either `userId`, `userName`, `email`, `aadGroupId` or `aadGroupName`

--8<-- "docs/cmd/_global.md"

## Remarks

For the `userId`, `userName`, `email`, `aadGroupId` or `aadGroupName` options you can specify multiple values by separating them with a comma. If one of the specified entries is not valid, the command will fail with an error message showing the list of invalid values.

## Examples

Add a user with the userName parameter to a SharePoint group with the groupId parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userName "Alex.Wilber@contoso.com"
```

Add multiple users with the userName parameter to a SharePoint group with the groupId parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userName "Alex.Wilber@contoso.com, Adele.Vance@contoso.com"
```

Add a user with the email parameter to a SharePoint group with the groupName parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Contoso Site Owners" --email "Alex.Wilber@contoso.com"
```

Add multiple users with the email parameter to a SharePoint group with the groupName parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupName "Contoso Site Owners" --email "Alex.Wilber@contoso.com, Adele.Vance@contoso.com"
```

Add a user with the userId parameter to a SharePoint group with the groupId parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userId 5
```

Add multiple users with the userId parameter to a SharePoint group with the groupId parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --userId "5,12"
```

Add multiple users with the aadUserId parameter to a SharePoint group with the groupId parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --aadUserId "56ca9023-3449-4e98-a96a-69e81a6f4983,a96c1755-8aec-4f4d-955f-08504510dc56"
```

Add multiple users with the aadUserName parameter to a SharePoint group with the groupId parameter

```sh
m365 spo group member add --webUrl https://contoso.sharepoint.com/sites/SiteA --groupId 5 --aadUserName "Azure group one, Azure group two"
```

## Response

=== "JSON"

    ```json
    [
      {
        "AllowedRoles": [
          0
        ],
        "CurrentRole": 0,
        "DisplayName": "John Doe",
        "Email": "john.doe@contoso.onmicrosoft.com",
        "InvitationLink": null,
        "IsUserKnown": true,
        "Message": null,
        "Status": true,
        "User": "i:0#.f|membership|john.doe@contoso.onmicrosoft.com"
      }
    ]
    ```

=== "Text"

    ```text
    DisplayName  Email
    -----------  ---------------------------------
    John Doe     john.doe@contoso.onmicrosoft.com
    ```

=== "CSV"

    ```csv
    DisplayName,Email
    John Doe,john.doe@contoso.onmicrosoft.com
    ```
