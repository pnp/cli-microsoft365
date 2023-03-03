# pa app export

Exports the specified Power App

## Usage

```sh
m365 pa app export [options]
```

## Options

`-i, --id <id>`
: The id of the Power App to export

`-e, --environment <environment>`
: The name of the environment for which to export the app

`-n, --packageDisplayName [packageDisplayName]`
: The display name to use in the exported package

`-d, --packageDescription [packageDescription]`
: The description to use in the exported package

`-c, --packageCreatedBy [packageCreatedBy]`
: The name of the person to be used as the creator of the exported package

`-s, --packageSourceEnvironment [packageSourceEnvironment]`
: The name of the source environment from which the exported package was taken

`-p, --path [path]`
: The path to save the exported package to

--8<-- "docs/cmd/_global.md"

## Examples

Export the specified Power App as a ZIP file

```sh
m365 pa app export --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --id 3989cb59-ce1a-4a5c-bb78-257c5c39381d --packageDisplayName "PowerApp" --packageDescription "Power App Description" --packageCreatedBy "John Doe" --packageSourceEnvironment "Contoso" --path "C:/Users/John/Documents"
```

## Response

=== "JSON"

    ```json
    ./PowerApp.zip
    ```

=== "Text"

    ```text
    ./PowerApp.zip
    ```

=== "CSV"

    ```csv
    ./PowerApp.zip
    ```

=== "Markdown"

    ```md
    ./PowerApp.zip
    ```
