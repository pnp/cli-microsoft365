# aad user set

Updates information of the specified user

## Usage

```sh
m365 aad user set [options]
```

## Options

`-i, --objectId [objectId]`
: The object ID of the user to update. Specify `objectId` or `userPrincipalName` but not both.

`-n, --userPrincipalName [userPrincipalName]`
: User principal name of the user to update. Specify `objectId` or `userPrincipalName` but not both.

`--accountEnabled [accountEnabled]`
: Boolean value specifying whether the account is enabled. Valid values are `true` or `false`.

`--resetPassword`
: If specified, the password of the user will be reset. This will make the parameter `newPassword` required.

`--forceChangePasswordNextSignIn`
: If specified, the user will have to change his password the next time they log in. Can only be set in combination with `resetPassword`.

`--forceChangePasswordNextSignInWithMfa`
: Whether the user should change his/her password on the next login and setup MFA. Can only be set in combination with `resetPassword`.

`--currentPassword [currentPassword]`
: Current password of the user that is signed in. If this parameter is set, `newPassword` is mandatory. Can't be combined with `resetPassword`.

`--newPassword [newPassword]`
: New password to be set. Must be set when specifying either `resetPassword` or `currentPassword`.

`--displayName [displayName]`
: The name to display in the address book for the user.

`--firstName [firstName]`
: The given name (first name) of the user. Maximum length is 64 characters.

`--lastName [lastName]`
: The user's surname (family name or last name). Maximum length is 64 characters.

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
: User ID of the user's manager. Specify `managerUserId`, `managerUserName` or `removeManger` but not both.

`--managerUserName [managerUserName]`
: User principal name of the manager. Specify `managerUserId`, `managerUserName` or `removeManger` but not both.

`--removeManger`
: Remove currently set manager. The user will have no manager when this flag is set. Specify `managerUserId`, `managerUserName` or `removeManger` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

This command allows using unknown options.

If the user with the specified id or user name doesn't exist, you will get a `Resource 'xyz' does not exist or one of its queried reference-property objects are not present.` error.

## Examples

Update specific property _department_ of user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --Department IT
```

Update multiple properties of user with name _steve@contoso.onmicrosoft.com_

```sh
m365 aad user set --userPrincipalName steve@contoso.onmicrosoft.com --CompanyName Contoso --firstName John --lastName Doe --jobTitle "Sales Manager" --companyName Contoso --department Sales --officeLocation "New York"
```

Enable user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --accountEnabled true
```

Disable user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --accountEnabled false
```

Enable user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --accountEnabled true
```

Reset password of a given user by userPrincipalName and require the user to change the password on the next sign in

```sh
m365 aad user set --userPrincipalName steve@contoso.onmicrosoft.com --resetPassword --newPassword 6NLUId79Lc24 --forceChangePasswordNextSignIn
```

Change password of the currently logged in user

```sh
m365 aad user set --objectId 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --currentPassword SLBF5gnRtyYc --newPassword 6NLUId79Lc24
```

Updates a user with a manager

```sh
m365 aad user set --displayName "John Doe" --userName "john.doe@contoso.com" --managerUserName "adele@contoso.com"
```

Updates a user by removing its manager

```sh
m365 aad user set --removeManger
```

## Response

The command won't return a response on success.
