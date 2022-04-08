# teams app remove

Removes a Teams app from the organization's app catalog

## Usage

```sh
m365 teams app remove [options]
```

## Options

`-i, --id <id>`
: ID of the Teams app to remove. Needs to be available in your organization\'s app catalog.

`--confirm`
: Don't prompt for confirming removing the app

--8<-- "docs/cmd/_global.md"

### Remarks

You can only remove a Teams app as a global administrator.

## Examples

Remove the Teams app with ID _83cece1e-938d-44a1-8b86-918cf6151957_ from the organization's app catalog. Will prompt for confirmation before actually removing the app.

```sh
m365 teams app remove --id 83cece1e-938d-44a1-8b86-918cf6151957
```

Remove the Teams app with ID _83cece1e-938d-44a1-8b86-918cf6151957_ from the organization's app catalog. Don't prompt for confirmation.

```sh
m365 teams app remove --id 83cece1e-938d-44a1-8b86-918cf6151957 --confirm
```
