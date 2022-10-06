# pp environment get

Gets information about the specified Power Platform environment

## Usage

```sh
m365 pp environment get [options]
```

## Options

`-n, --name <name>`
: The name of the environment to get information about

`-a, --asAdmin`
: Run the command as admin and retrieve details of environments you do not have explicitly assigned permissions to

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.
    Register CLI for Microsoft 365 or Azure AD application as a management application for the Power Platform using 
    m365 pp managementapp add [options] 

## Examples

Get information about the Power Platform environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_

```sh
m365 pp environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```

Get information about the Power Platform environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_ as Admin

```sh
m365 pp environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --asAdmin
```
