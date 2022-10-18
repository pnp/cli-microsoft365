# pp solution list

Lists a specific solution in a given environment.

## Usage

```sh
m365 pp solution get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment

<<<<<<< HEAD
`-i --id`
: The id of the card

`-n, --name`
=======
`-n, --name <name>`
>>>>>>> 9881501e (solution-get)
: The unique name of the card

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

<<<<<<< HEAD
List a specific solution in a specific environment based on the name
=======
List a specific solution in a specific environment
>>>>>>> 9881501e (solution-get)

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "Default"
```

<<<<<<< HEAD
List a specific solution in a specific environment based on the name as Admin
=======
List a specific solution in a specific environment as Admin
>>>>>>> 9881501e (solution-get)

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "Default" --asAdmin
```
<<<<<<< HEAD

List a specific solution in a specific environment based on the id

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --id "ee62fd63-e49e-4c09-80de-8fae1b9a427e"
```

List a specific solution in a specific environment based on the id as Admin

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --id "ee62fd63-e49e-4c09-80de-8fae1b9a427e" --asAdmin
```
=======
>>>>>>> 9881501e (solution-get)
