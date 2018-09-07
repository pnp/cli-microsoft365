# aad sp get

Gets information about the specific service principal

## Usage

```sh
aad sp get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --appId [appId]`|ID of the application for which the service principal should be retrieved
`-n, --displayName [displayName]`|Display name of the application for which the service principal should be retrieved
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to Azure Active Directory Graph, using the [aad login](../login.md) command.

## Remarks

To get information about a service principal, you have to first log in to Azure Active Directory Graph using the [aad login](../login.md) command, eg. `aad login`.

When looking up information about a service principal you should specify either its `appId` or `displayName` but not both. If you specify both values, the command will fail with an error.

## Examples

Return details about the service principal with appId _b2307a39-e878-458b-bc90-03bc578531d6_.

```sh
aad sp get --appId b2307a39-e878-458b-bc90-03bc578531d6
```

Return details about the _Microsoft Graph_ service principal.

```sh
aad sp get --displayName "Microsoft Graph"
```

## More information

- Application and service principal objects in Azure Active Directory (Azure AD): [https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects)