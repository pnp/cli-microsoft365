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

This command will execute the following steps.

- Stop the Microsoft Teams client. This will kill all the running `Teams.exe` tasks.

- Clear the Microsoft Teams cached files. For Windows it will delete all files and folders in the %appdata%\Microsoft\Teams directory. For macOS it will delete all files and folders in the  ~/Library/Application Support/Microsoft/Teams directory.

!!! important
    - You won't lose any user data by clearing the cache.
    - Restarting Teams after you clear the cache might take longer than usual because the Teams cache files have to be rebuilt.

If you run the command in the CLI Docker container, you will get the following error message:

> Because you're running CLI for Microsoft 365 in a Docker container, we can't clear the cache on your host. Instead run this command on your host using 'npx ...'.

The command works only on Windows and macOS. If you run it on a different operating system, you will get the `'abc' platform is unsupported for this command` error.

## Examples

Removes the Microsoft Teams client cache

```sh
m365 teams cache remove
```

## More information

- [Clear the Teams client cache guidance](https://docs.microsoft.com/microsoftteams/troubleshoot/teams-administration/clear-teams-cache)