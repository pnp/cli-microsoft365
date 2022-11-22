# pp solution publisher add

Adds a specified publisher in a given environment.

## Usage

```sh
m365 pp solution publisher add [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-n, --name [name]`
: The unique name of the publisher.

`--displayName <displayName>`
: The friendly name of the publisher

`--prefix <prefix>`
: The prefix of the publisher

`--choiceValuePrefix <choiceValuePrefix>`
: The choice value prefix of the publisher

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

Add a specific publisher in a specific environment

```sh
m365 pp solution publisher add --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "Contoso" --displayName "Contoso" --prefix "new" --choiceValuePrefix 10000
```

Adds a specific publisher in a specific environment as Admin

```sh
m365 pp solution publisher add --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "Contoso" --displayName "Contoso" --prefix "new" --choiceValuePrefix 10000 --asAdmin
```

## Response

The command won't return a response on success.
