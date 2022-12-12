# pp solution publisher remove

Removes the specified publisher in the specified Power Platform environment

## Usage

```sh
m365 pp solution publisher remove [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-i, --id [id]`
: The id of the solution. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The name of the solution. Specify either `id` or `name` but not both.

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples

Removes the specified publisher based on the name parameter

```sh
m365 pp solution publisher remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name "Publisher Name"
```

Removes the specified publisher based on the name parameter without confirmation

```sh
m365 pp solution publisher remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name "Publisher Name" --confirm
```

Removes the specified publisher based on the name parameter as admin

```sh
m365 pp solution publisher remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name "Publisher Name" --asAdmin
```

Removes the specified publisher owned by the currently signed-in user based on the id parameter

```sh
m365 pp solution publisher remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 00000001-0000-0000-0001-00000000009b
```

## Response

The command won't return a response on success.
