# aad approleassignment list

Lists AppRoleAssignments for the specified application registration

## Usage

```sh
aad approleassignment list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --appId <appId>`|Application (client) Id of the App Registration for which the configured appRoles should be retrieved
`-n, --displayName <displayName>`|Display name of the application for which the configured appRoles should be retrieved
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Specify either the appId or displayName but not both. If you specify both values, the command will fail with an error.

## Examples

List AppRoles assigned to service principal with Application (client) ID _b2307a39-e878-458b-bc90-03bc578531d6_.

```sh
aad approleassignment list --appId b2307a39-e878-458b-bc90-03bc578531d6
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)