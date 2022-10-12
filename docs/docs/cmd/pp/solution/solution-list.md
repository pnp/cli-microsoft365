# pp solution list

Lists solutions in a given environment.

## Usage

```sh
m365 pp solution list [options]
```

## Options

`-e, --environment <environment>`
The name of the environment to list the solutions for

`-a, --asAdmin`
Set, to retrieve the solutions as admin for environments you are not a member of.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

List all solutions for the given environment

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339"
```

List all solutions for the given environment as Admin

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --asAdmin
```
