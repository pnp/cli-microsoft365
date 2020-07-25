# flow remove

Removes the specified Microsoft Flow

## Usage

```sh
m365 flow remove [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: The name of the Microsoft Flow to remove

`-e, --environment <environment>`
: The name of the environment to which the Flow belongs

`--asAdmin`
: Set, to remove the Flow as admin

`--confirm`
: Don't prompt for confirmation

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

By default, the command will try to remove a Microsoft Flow you own. If you want to remove a Flow owned by another user, use the `asAdmin` flag.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Microsoft Flow with the name you specified doesn't exist, you will get the `Error: Resource 'abc' does not exist in environment 'xyz'` error.

## Examples

Removes the specified Microsoft Flow owned by the currently signed-in user

```sh
m365 flow remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Removes the specified Microsoft Flow owned by the currently signed-in user without confirmation

```sh
m365 flow remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --confirm
```

Removes the specified Microsoft Flow owned by another user

```sh
m365 flow remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
```

Removes the specified Microsoft Flow owned by another user without confirmation

```sh
m365 flow remove --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin --confirm
```
