# spo user list

Lists all the users within specific web

## Usage

```sh
m365 spo user list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the web to list the users from

--8<-- "docs/cmd/_global.md"

## Examples

Get list of users in web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo user list --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    [
      {
        "Id": 10,
        "IsHiddenInUI": false,
        "LoginName": "i:0#.f|membership|johndoe@contoso.onmicrosoft.com",
        "Title": "John DOe",
        "PrincipalType": 1,
        "Email": "johndoe@contoso.onmicrosoft.com",
        "Expiration": "",
        "IsEmailAuthenticationGuestUser": false,
        "IsShareByEmailGuestUser": false,
        "IsSiteAdmin": false,
        "UserId": {
          "NameId": "100320022ec308a7",
          "NameIdIssuer": "urn:federation:microsoftonline"
        },
        "UserPrincipalName": "johndoe@contoso.onmicrosoft.com"
      }
    ]
    ```

=== "Text"

    ```text
    Id          Title                           LoginName
    ----------  ------------------------------  --------------------------------------------------------------------------
    10          John Doe                        i:0#.f|membership|johndoe@contoso.onmicrosoft.com
    ```

=== "CSV"

    ```csv
    Id,Title,LoginName
    10,John Doe,i:0#.f|membership|johndoe@contoso.onmicrosoft.com
    ```

=== "Markdown"

    ```md
    # spo user list --webUrl "https://contoso.sharepoint.com/sites/project-x"

    Date: 4/10/2023

    ## John Doe (10)

    Property | Value
    ---------|-------
    Id | 7
    IsHiddenInUI | false
    LoginName | c:0o.c\|membership\|johndoe@contoso.onmicrosoft.com
    Title | John Doe
    PrincipalType | 1
    Email | johndoe@contoso.onmicrosoft.com
    Expiration | 
    IsEmailAuthenticationGuestUser | false
    IsShareByEmailGuestUser | false
    IsSiteAdmin | false
    UserId | {"NameId":"1003200107a7bdf7","NameIdIssuer":"urn:federation:microsoftonline"}
    UserPrincipalName | johndoe@contoso.onmicrosoft.com
    ```
