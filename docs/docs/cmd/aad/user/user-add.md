# aad user add

Creates a new user

## Usage

```sh
m365 aad user add [options]
```

## Options

`--displayName <displayName>`
: The name to display in the address book for the user.

`--userName <userName>`
: The user principal name (someuser@contoso.com).

`--accountEnabled [accountEnabled]`
: Whether the account is enabled. Possible values: `true`, `false`. Default value is `true`.

`--mailNickname [mailNickname]`
: The mail alias for the user. By default this value will be extracted from `userName`.

`--password [password]`
: The password for the user. When not specified, a password will be generated.

`--firstName [firstName]`
: The given name (first name) of the user. Maximum length is 64 characters.

`--lastName [lastName]`
: The user's surname (family name or last name). Maximum length is 64 characters.

`--forceChangePasswordNextSignIn`
: Whether the user should change his/her password on the next login.

`--forceChangePasswordNextSignInWithMfa`
: Whether the user should change his/her password on the next login and setup MFA.

`--usageLocation [usageLocation]`
: A two letter [country code](https://learn.microsoft.com/en-us/partner-center/commercial-marketplace-co-sell-location-codes#country-and-region-codes) (ISO standard 3166). Required for users that will be assigned licenses.

`--officeLocation [officeLocation]` 
: The office location in the user's place of business.

`--jobTitle [jobTitle]`
: The user's job title. Maximum length is 128 characters.

`--companyName [companyName]`
: The company name which the user is associated. The maximum length is 64 characters.

`--department [department]`
: The name for the department in which the user works. Maximum length is 64 characters.

`--preferredLanguage [preferredLanguage]`
: The preferred language for the user. Should follow [ISO 639-1 Code](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a). Example: `en-US`.

`--managerUserId [managerUserId]`
: User ID of the user's manager. Specify `managerUserId` or `managerUserName` but not both.

`--managerUserName [managerUserName]`
: User principal name of the manager. Specify `managerUserId` or `managerUserName` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global administrator, User administrator or Privileged Authentication administrator

!!! note
    After running this command, it may take a minute before the user is effectively created in the tenant.

This command allows using unknown options. For a comprehensive list of user properties, please refer to the [Graph documentation page](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0#properties).

If the specified option is not found, you will receive a `Resource 'xyz' does not exist or one of its queried reference-property objects are not present.` error.

## Examples

Create a user and let him/her update the password at first login

```sh
m365 aad user add --displayName "John Doe" --userName "john.doe@contoso.com" --password "SomePassw0rd" --forceChangePasswordNextSignIn
```

Create a user with job information

```sh
m365 aad user add --displayName "John Doe" --userName "john.doe@contoso.com" --password "SomePassw0rd" --firstName John --lastName Doe --jobTitle "Sales Manager" --companyName Contoso --department Sales --officeLocation Vosselaar --forceChangePasswordNextSignIn
```

Create a user with language information

```sh
m365 aad user add --displayName "John Doe" --userName "john.doe@contoso.com" --preferredLanguage "nl-BE" --usageLocation BE --forceChangePasswordNextSignIn
```

Create a user with a manager

```sh
m365 aad user add --displayName "John Doe" --userName "john.doe@contoso.com" --managerUserName "adele@contoso.com"
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
      "officeLocation": "Vosselaar",
      "preferredLanguage": "nl-BE",
      "surname": "Doe",
      "userPrincipalName": "john.doe@contoso.com",
      "password": "SomePassw0rd"
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
    officeLocation   : Vosselaar
    password         : SomePassw0rd
    preferredLanguage: nl-BE
    surname          : Doe
    userPrincipalName: john.doe@contoso.com
    ```

=== "CSV"

    ```csv
    id,businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName,password
    990e2425-f595-43bc-85ed-b89a44093793,[],John Doe,John,Sales Manager,,,Vosselaar,nl-BE,Doe,john.doe@contoso.com,SomePassw0rd
    ```

=== "Markdown"

    ```md
    # aad user add --displayName "John Doe" --userName "john.doe@contoso.com" --password "SomePassw0rd" --firstName "John" --lastName "Doe" --jobTitle "Sales Manager" --officeLocation "Vosselaar" --preferredLanguage "nl-BE"

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
    officeLocation | Vosselaar
    preferredLanguage | nl-BE
    surname | Doe
    userPrincipalName | john.doe@contoso.com
    password | SomePassw0rd
    ```
