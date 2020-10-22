# aad approleassignment list

Lists app role assignments for the specified application registration

## Usage

```sh
m365 aad approleassignment list [options]
```

## Options

`-i, --appId [appId]`
: Application (client) Id of the App Registration for which the configured app roles should be retrieved

`-n, --displayName [displayName]`
: Display name of the application for which the configured app roles should be retrieved

`--objectId [objectId]`
: ObjectId of the application for which the configured app roles should be retrieved

--8<-- "docs/cmd/_global.md"

## Remarks

Specify either the `appId`, `objectId` or `displayName`. If you specify more than one option value, the command will fail with an error.

## Examples

List app roles assigned to service principal with Application (client) ID _b2307a39-e878-458b-bc90-03bc578531d6_.

```sh
m365 aad approleassignment list --appId b2307a39-e878-458b-bc90-03bc578531d6
```

List app roles assigned to service principal with Application display name _MyAppName_.

```sh
m365 aad approleassignment list --displayName 'MyAppName'
```

List app roles assigned to service principal with ObjectId _b2307a39-e878-458b-bc90-03bc578531dd_.

```sh
m365 aad approleassignment list --objectId b2307a39-e878-458b-bc90-03bc578531dd
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)