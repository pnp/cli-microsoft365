# aad user recyclebinitem list

Lists users from the recycle bin in the current tenant

## Usage

```sh
m365 aad user recyclebinitem list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

List all removed users

```sh
m365 aad user recyclebinitem list
```

## Response

=== "JSON"

    ```json
    [
      {
        "businessPhones": [],
        "displayName": "John Doe",
        "givenName": "John Doe",
        "jobTitle": "Developer",
        "mail": "john@contoso.com",
        "mobilePhone": "0476345130",
        "officeLocation": "Washington",
        "preferredLanguage": "nl-BE",
        "surname": "John",
        "userPrincipalName": "7e06b56615f340138bf879874d52e68a277@contoso.com",
        "id": "7e06b566-15f3-4013-8bf8-79874d52e68a"
      }
    ]
    ```

=== "Text"

    ```text
    id                                    displayName  userPrincipalName
    ------------------------------------  -----------  -----------------------------------------------
    7e06b566-15f3-4013-8bf8-79874d52e68a  John Doe     7e06b56615f340138bf879874d52e68a277@contoso.com
    ```

=== "CSV"

    ```csv
    id,displayName,userPrincipalName
    7e06b566-15f3-4013-8bf8-79874d52e68a,John Doe,7e06b56615f340138bf879874d52e68a277@contoso.com
    ```

=== "Markdown"

    ```md
    # aad user recyclebinitem list

    Date: 14/02/2023

    ## John Doe (7e06b566-15f3-4013-8bf8-79874d52e68a)

    Property | Value
    ---------|-------
    businessPhones | []
    displayName | John Doe
    givenName | John Doe
    jobTitle | Developer
    mail | john@contoso.com
    mobilePhone | 0476345130
    officeLocation | Washington
    preferredLanguage | nl-BE
    surname | John
    userPrincipalName | 7e06b56615f340138bf879874d52e68a277@contoso.com
    id | 7e06b566-15f3-4013-8bf8-79874d52e68a
    ```
