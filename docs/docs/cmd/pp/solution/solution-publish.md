# pp solution publish

Publishes the components of a specified solution in the specified Power Platform environment

## Usage

```sh
m365 pp solution publish [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-i, --id [id]`
: The id of the solution. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The unique name (not the display name) of the solution. Specify either `id` or `name` but not both.

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples

Publishes the components of a specified solution with a specific name, owned by the currently signed-in user

```sh
m365 pp solution publish --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name "Solution Name"
```

Publishes the components of a specified solution owned by the currently signed-in user based on the id parameter and waits for completion

```sh
m365 pp solution publish --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 00000001-0000-0000-0001-00000000009b --wait
```

Publishes the components of a specified solution owned by another user based on the name parameter

```sh
m365 pp solution publish --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name "Solution Name" --asAdmin
```

## Response

The command won't return a response on success.
