# teams app update

Updates Teams app in the organization's app catalog

## Usage

```sh
m365 teams app update [options]
```

## Options

`-i, --id [id]`
: ID of the app to update. Specify either id or name, but not both

`-n, --name [name]`
: The display name of the app to update. Specify either id or name, but not both

`-p, --filePath <filePath>`
: Absolute or relative path to the Teams manifest zip file to update in the app catalog

--8<-- "docs/cmd/_global.md"

## Remarks

You can only update a Teams app as a global administrator.

## Examples

Update the Teams app with ID _83cece1e-938d-44a1-8b86-918cf6151957_ from file _teams-manifest.zip_

```sh
m365 teams app update --id 83cece1e-938d-44a1-8b86-918cf6151957 --filePath ./teams-manifest.zip
```

Update the Teams app with name _Test app_ from file _teams-manifest.zip_

```sh
m365 teams app update --name "Test app" --filePath ./teams-manifest.zip
```
