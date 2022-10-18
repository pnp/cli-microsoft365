# pp solution list

Lists a specific solution in a given environment.

## Usage

```sh
m365 pp solution get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment

`-n, --name <name>`
: The unique name of the card

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List a specific solution in a specific environment

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "Default"
```

List a specific solution in a specific environment as Admin

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "Default" --asAdmin
```
