# pp solution get

Gets a specific solution in a given environment.

## Usage

```sh
m365 pp solution get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment

`-i --id [id]`
: The ID of the solution. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The unique name of the solution, not the friendly name. Specify either `id` or `name` but not both.

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List a specific solution in a specific environment based on the name

```sh
m365 pp solution get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "Default"
```

List a specific solution in a specific environment based on the name as Admin

```sh
m365 pp solution get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "Default" --asAdmin
```

List a specific solution in a specific environment based on the id

```sh
m365 pp solution get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --id "ee62fd63-e49e-4c09-80de-8fae1b9a427e"
```

List a specific solution in a specific environment based on the id as Admin

```sh
m365 pp solution get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --id "ee62fd63-e49e-4c09-80de-8fae1b9a427e" --asAdmin
```
