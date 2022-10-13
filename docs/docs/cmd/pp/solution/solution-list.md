# pp solution list

Lists solutions in a given environment.

## Usage

```sh
m365 pp solution list [options]
```

## Options

`-e, --environment <environment>`
The name of the environment

`-a, --asAdmin`
Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List all solutions in a specific environment

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339"
```

List all solutions in a specific environment as Admin

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --asAdmin
```
