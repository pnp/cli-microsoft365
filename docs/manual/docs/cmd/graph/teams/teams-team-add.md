# graph teams add

Add a Microsoft Teams team

## Usage

```sh
graph teams add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --groupId [groupId]`|The ID of the O365 group to add a Microsoft Teams team
`-n, --name [name]`|Microsoft Teams team name
`-d, --description  [description]`|The description of the Microsoft Teams team
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To add a Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Add a new Microsoft Teams team by creating a group 

```sh
graph teams add --name 'Architecture' --description 'Architecture Discussion'
```

Add a new Microsoft Teams team for a group  

```sh
graph teams add --groupId 6d551ed5-a606-4e7d-b5d7-36063ce562cc
```
