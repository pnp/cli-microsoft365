# spo user ensure

Ensures that a user is available on a specific site

## Usage

```sh
m365 spo user ensure [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site.

`--aadId [--aadId]`
: Id of the user in Azure AD. Specify either `aadId` or `userName` but not both.

`--userName [userName]`
: User's UPN (user principal name, e.g. john@contoso.com). Specify either `aadId` or `userName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Ensures a user by its Azure AD Id

```sh
m365 spo user ensure --webUrl https://contoso.sharepoint.com/sites/project --aadId e254750a-eaa4-44f6-9517-b74f65cdb747
```

Ensures a user by its user principal name

```sh
m365 spo user ensure --webUrl https://contoso.sharepoint.com/sites/project --userName john@contoso.com
```

## Response

=== "JSON"

    ```json
    {
      "Id": 35,
      "IsHiddenInUI": false,
      "LoginName": "i:0#.f|membership|john@contoso.com",
      "Title": "John Doe",
      "PrincipalType": 1,
      "Email": "john@contoso.com",
      "Expiration": "",
      "IsEmailAuthenticationGuestUser": false,
      "IsShareByEmailGuestUser": false,
      "IsSiteAdmin": false,
      "UserId": {
        "NameId": "1003200274f51d2d",
        "NameIdIssuer": "urn:federation:microsoftonline"
      },
      "UserPrincipalName": "john@contoso.com"
    }
    ```

=== "Text"

    ```text
    Email                         : john@contoso.com
    Expiration                    :
    Id                            : 35
    IsEmailAuthenticationGuestUser: false
    IsHiddenInUI                  : false
    IsShareByEmailGuestUser       : false
    IsSiteAdmin                   : false
    LoginName                     : i:0#.f|membership|john@contoso.com
    PrincipalType                 : 1
    Title                         : John Doe
    UserId                        : {"NameId":"1003200274f51d2d","NameIdIssuer":"urn:federation:microsoftonline"}
    UserPrincipalName             : john@contoso.com
    ```

=== "CSV"

    ```csv
    Id,IsHiddenInUI,LoginName,Title,PrincipalType,Email,Expiration,IsEmailAuthenticationGuestUser,IsShareByEmailGuestUser,IsSiteAdmin,UserId,UserPrincipalName
    35,,i:0#.f|membership|john@contoso.com,John Doe,1,john@contoso.com,,,,,"{""NameId"":""100320009d80e5de"",""NameIdIssuer"":""urn:federation:microsoftonline""}",john@contoso.com
    ```

=== "Markdown"

    ```md
    # spo user ensure --webUrl "https://mathijsdev2.sharepoint.com" --userName "john@contoso.com"

    Date: 18/02/2023

    ## John Doe (35)

    Property | Value
    ---------|-------
    Id | 35
    IsHiddenInUI | false
    LoginName | i:0#.f\|membership\|john@contoso.com
    Title | John Doe
    PrincipalType | 1
    Email | john@contoso.com
    Expiration |
    IsEmailAuthenticationGuestUser | false
    IsShareByEmailGuestUser | false
    IsSiteAdmin | false
    UserId | {"NameId":"100320009d80e5de","NameIdIssuer":"urn:federation:microsoftonline"}
    UserPrincipalName | john@contoso.com
    ```
