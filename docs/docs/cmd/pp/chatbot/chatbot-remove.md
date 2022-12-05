# pp chatbot remove

Removes the specified chatbot

## Usage

```sh
m365 pp chatbot remove [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-i, --id [id]`
: The id of the chatbot. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The name of the chatbot. Specify either `id` or `name` but not both.

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes the specified Microsoft Power Platform chatbot owned by the currently signed-in user based on the name parameter

```sh
m365 pp chatbot remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name "Chatbot Name"
```

Removes the specified Microsoft Power Platform chatbot owned by the currently signed-in user based on the id parameter without confirmation

```sh
m365 pp chatbot remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 9d9a13d0-6255-ed11-bba2-000d3adf774e --confirm
```

Removes the specified Microsoft Power Platform chatbot owned by another user based on the name parameter

```sh
m365 pp chatbot remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name "Chatbot Name" --asAdmin
```

## Response

The command won't return a response on success.
