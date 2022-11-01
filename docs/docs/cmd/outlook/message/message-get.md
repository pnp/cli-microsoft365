# outlook message get

Retrieves specified message

## Usage

```sh
m365 outlook message get [options]
```

## Options

`-i, --id <id>`
: ID of the message

`--userId [userId]`
: ID of the user from which to retrieve the message. Specify either `userId` or `userPrincipalName`, but not both. This option is required when using application permissions.

`--userPrincipalName [userPrincipalName]`
: UPN of the user from which to retrieve the message Specify either `userId` or `userPrincipalName`, but not both. This option is required when using application permissions.

--8<-- "docs/cmd/_global.md"

## Examples

Get a specific message using delegated permissions

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA=
```

Get a specific message using delegated permissions from a shared mailbox

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA= --userPrincipalName sharedmailbox@tenant.com
```

Get a specific message from a specific user retrieved by user ID using application permissions

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA= --userId 6799fd1a-723b-4eb7-8e52-41ae530274ca
```

Get a specific message from a specific user retrieved by user principal name using application permissions

```sh
m365 outlook message get --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA= --userPrincipalName user@tenant.com
```
