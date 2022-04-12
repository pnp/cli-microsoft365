# teams cache remove

Removes the Microsoft Teams client cache

## Usage

```sh
m365 teams cache remove [options]
```

## Options

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Remarks

!!! note
    - You won't lose any user data by clearing the cache.
    - Restarting Teams after you clear the cache might take longer than usual because the Teams cache files have to be rebuilt.

If the command is executed within the CLI Docker container, you will get the `Because you're running CLI for Microsoft 365 in a Docker container, we can't clear the cache on your host. Instead run this command on your host using 'npx ...'` error.

If the command isn't executed from a Windows or MacOS system, you will get the `'abc' platform is unsupported for this command` error.

## Examples

Removes the Microsoft Teams client cache

```sh
m365 teams cache remove
```

## More information

- Guidance from the Microsoft Docs article: [Clear the Teams client cache](https://docs.microsoft.com/en-us/microsoftteams/troubleshoot/teams-administration/clear-teams-cache)