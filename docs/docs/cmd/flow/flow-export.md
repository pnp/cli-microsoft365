# flow export

Exports the specified Power Automate flow

## Usage

```sh
m365 flow export [options]
```

## Options

`-i, --id <id>`
: The id of the Power Automate flow to export

`-e, --environmentNameName <environmentName>`
: The name of the environment for which to export the flow

`-n, --packageDisplayName [packageDisplayName]`
: The display name to use in the exported package

`-d, --packageDescription [packageDescription]`
: The description to use in the exported package

`-c, --packageCreatedBy [packageCreatedBy]`
: The name of the person to be used as the creator of the exported package

`-s, --packageSourceEnvironment [packageSourceEnvironment]`
: The name of the source environment from which the exported package was taken

`-f, --format [format]`
: Export format type. `json,zip`. Default `zip`

`-p, --path [path]`
: The path to save the exported package to

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Power Automate flow with the id you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error.

## Examples

Export the specified Power Automate flow as a ZIP file

```sh
m365 flow export --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Export the specified Power Automate flow as a JSON file

```sh
m365 flow export --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --format json
```

Export the specified Power Automate flow as a ZIP file, specifying a Display Name of 'My flow name' to be embedded into the package

```sh
m365 flow export --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --packageDisplayName 'My flow name'
```

Export the specified Power Automate flow as a ZIP file with the filename 'MyFlow.zip' saved to the current directory

```sh
m365 flow export --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --path './MyFlow.zip'
```

## Response

When the specified Power Automate flow is exported as a ZIP file

```text
"./20221117t000000z816608b996494b328cf9f9794f160209_20221117124608.zip"
```

When the specified Power Automate flow is exported as a JSON file

```text
"./MyFlow.json"
```

When the specified Power Automate flow is exported by specifying a package display Name

```text
"./Myflowname_20221117125033.zip"
```
