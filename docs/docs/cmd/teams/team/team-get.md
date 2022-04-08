# teams team get

Gets information about the specified Microsoft Teams team

## Usage

```sh
m365 teams team get
```

## Options

`-i, --id [id]`
: The ID of the Microsoft Teams team to retrieve information for. Specify either id or name but not both

`-n, --name [name]`
: The display name of the Microsoft Teams team to retrieve information for. Specify either id or name but not both

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the Microsoft Teams team with id _2eaf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
m365 teams team get --id 2eaf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about Microsoft Teams team with name _Team Name_

```sh
m365 teams team get --name "Team Name"
```
