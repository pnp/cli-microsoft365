# pp solution publisher get

Get information about the specified publisher in a given environment.

## Usage

```sh
m365 pp solution publisher get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-i --id [id]`
: The ID of the solution. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The unique name (not the display name) of the solution. Specify either `id` or `name` but not both.

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

Gets a specific publisher in a specific environment based on name

```sh
m365 pp solution publisher get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "MicrosoftCorporation"
```

Gets a specific publisher in a specific environment based on name as Admin

```sh
m365 pp solution publisher get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "MicrosoftCorporation" --asAdmin
```

Gets a specific publisher in a specific environment based on id

```sh
m365 pp solution publisher get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --id "ee62fd63-e49e-4c09-80de-8fae1b9a427e"
```

Gets a specific publisher in a specific environment based on id as Admin

```sh
m365 pp solution publisher get --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --id "ee62fd63-e49e-4c09-80de-8fae1b9a427e" --asAdmin
```
