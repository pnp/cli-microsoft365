# aad user license add

Assigns a license to a user

## Usage

```sh
m365 aad user license add [options]
```

## Options

`--userId [userId]`
: The ID of the user. Specify either `userId` or `userName` but not both.

`--userName [userName]`
: User principal name of the user. Specify either `userId` or `userName` but not both.

`--ids <ids>`
: A comma separated list of IDs that specify the licenses to add.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    The user must have a `usageLocation` value in order to assign a license to it.

## Examples

Assign specific licenses to a specific user by UPN

```sh
m365 user license add --userName "john.doe@contoso.com" --ids "45715bb8-13f9-4bf6-927f-ef96c102d394,bea13e0c-3828-4daa-a392-28af7ff61a0f"
```

Assign specific licenses from a specific user by ID

```sh
m365 user license add --userId "5c241023-2ba5-4ea8-a516-a2481a3e6c51" --ids "45715bb8-13f9-4bf6-927f-ef96c102d394,bea13e0c-3828-4daa-a392-28af7ff61a0f"
```

## Response

=== "JSON"

    ```json
    {
      "businessPhones": [],
      "displayName": "John Doe",
      "givenName": null,
      "jobTitle": null,
      "mail": "John@contoso.onmicrosoft.com",
      "mobilePhone": null,
      "officeLocation": null,
      "preferredLanguage": null,
      "surname": null,
      "userPrincipalName": "John@contoso.onmicrosoft.com",
      "id": "eb77fbcf-6fe8-458b-985d-1747284793bc"
    }
    ```

=== "Text"

    ```text
    businessPhones   : []
    displayName      : John Doe
    givenName        : null
    id               : eb77fbcf-6fe8-458b-985d-1747284793bc
    jobTitle         : null
    mail             : John@contoso.onmicrosoft.com
    mobilePhone      : null
    officeLocation   : null
    preferredLanguage: null
    surname          : null
    userPrincipalName: John@contoso.onmicrosoft.com
    ```

=== "CSV"

    ```csv
    businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,id
    [],John Doe,,,John@contoso.onmicrosoft.com,,,,,John@contoso.onmicrosoft.com,eb77fbcf-6fe8-458b-985d-1747284793bc
    ```

=== "Markdown"

    ```md
    # aad user license add --userName "John@contoso.onmicrosoft.com" --ids "f30db892-07e9-47e9-837c-80727f46fd3d,606b54a9-78d8-4298-ad8b-df6ef4481c80"

    Date: 16/2/2023

    ## John Doe (eb77fbcf-6fe8-458b-985d-1747284793bc)

    Property | Value
    ---------|-------
    businessPhones | []
    displayName | John Doe
    givenName | null
    jobTitle | null
    mail | John@contoso.onmicrosoft.com
    mobilePhone | null
    officeLocation | null
    preferredLanguage | null
    surname | null
    userPrincipalName | John@contoso.onmicrosoft.com
    id | eb77fbcf-6fe8-458b-985d-1747284793bc
    ```
