# aad user recyclebinitem restore

Restores a user from the recycle bin in the current tenant

## Usage

```sh
m365 aad user recyclebinitem restore [options]
```

## Options

`--id <id>`
: ID of the deleted user.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global administrator, User administrator or Privileged Authentication administrator

!!! note
    After running this command, it may take a minute before the user will reappear in the active users.

## Examples

Restore user from the recycle bin

```sh
m365 aad user recyclebinitem restore --id 990e2425-f595-43bc-85ed-b89a44093793
```

## Response

=== "JSON"

    ```json
    {
      "id": "990e2425-f595-43bc-85ed-b89a44093793",
      "businessPhones": [],
      "displayName": "John Doe",
      "givenName": "John",
      "jobTitle": "Sales Manager",
      "mail": null,
      "mobilePhone": null,
      "officeLocation": null,
      "preferredLanguage": "nl-BE",
      "surname": "Doe",
      "userPrincipalName": "john.doe@contoso.com",
    }
    ```

=== "Text"

    ```text
    businessPhones   : []
    displayName      : John Doe
    givenName        : John
    id               : 990e2425-f595-43bc-85ed-b89a44093793
    jobTitle         : Sales Manager
    mail             : null
    mobilePhone      : null
    officeLocation   : null
    preferredLanguage: nl-BE
    surname          : Doe
    userPrincipalName: john.doe@contoso.com
    ```

=== "CSV"

    ```csv
    id,businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName
    990e2425-f595-43bc-85ed-b89a44093793,[],John Doe,John,Sales Manager,,,,nl-BE,Doe,john.doe@contoso.com
    ```

=== "Markdown"

    ```md
    # user recyclebin restore --id 990e2425-f595-43bc-85ed-b89a44093793

    Date: 16/02/2023

    ## John Doe (990e2425-f595-43bc-85ed-b89a44093793)

    Property | Value
    ---------|-------
    id | 990e2425-f595-43bc-85ed-b89a44093793
    businessPhones | []
    displayName | John Doe
    givenName | John
    jobTitle | Sales Manager
    mail | null
    mobilePhone | null
    officeLocation | null
    preferredLanguage | nl-BE
    surname | Doe
    userPrincipalName | john.doe@contoso.com
    ```
