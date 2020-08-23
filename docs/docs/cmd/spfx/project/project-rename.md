# spfx project rename

Renames SharePoint Framework project

## Usage

```sh
spfx project rename [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --newName <newName>`|New name for the project
`--generateNewId`|Generate a new solution ID for the project
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text|md`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Run this command in the folder where the project that you want to rename is located.

## Remarks

This command will update the project name in: _package.json_, _.yo-rc.json_, _package-solution.json_, _deploy-azure-storage.json_ and _README.md_.

## Examples

Renames SharePoint Framework project to contoso

```sh
spfx project rename --newName contoso
```

Renames SharePoint Framework project to contoso with new solution ID

```sh
spfx project rename --newName contoso --generateNewId
```
