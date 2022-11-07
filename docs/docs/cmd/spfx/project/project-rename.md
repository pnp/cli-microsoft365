# spfx project rename

Renames SharePoint Framework project

## Usage

```sh
m365 spfx project rename [options]
```

## Options

`-n, --newName <newName>`
: New name for the project

`--generateNewId`
: Generate a new solution ID for the project

--8<-- "docs/cmd/_global.md"

!!! important
    Run this command in the folder where the project that you want to rename is located.

## Remarks

This command will update the project name in: _package.json_, _.yo-rc.json_, _package-solution.json_, _deploy-azure-storage.json_ and _README.md_.

## Examples

Renames SharePoint Framework project to contoso

```sh
m365 spfx project rename --newName contoso
```

Renames SharePoint Framework project to contoso with new solution ID

```sh
m365 spfx project rename --newName contoso --generateNewId
```

## Response

The command won't return a response on success.
