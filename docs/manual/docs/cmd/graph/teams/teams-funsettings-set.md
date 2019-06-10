# graph teams funsettings set

Updates fun settings of a Microsoft Teams team

## Usage

```sh
graph teams funsettings set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Teams team for which to update settings
`--allowGiphy [allowGiphy]`|Set to `true` to allow giphy and to `false` to disable it
`--giphyContentRating [giphyContentRating]`|Settings to set content rating for giphy. Allowed values `Strict|Moderate`
`--allowStickersAndMemes [allowStickersAndMemes]`|Set to `true` to allow stickers and memes and to `false` to disable them
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To update fun settings of the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Allow giphy usage within a given Microsoft Teams team, setting the content rating for giphy to Moderate

```sh
graph teams funsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowGiphy true --giphyContentRating Moderate
```

Disallow usage of giphy within a given Microsoft Teams team

```sh
graph teams funsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowGiphy false
```

Disallow usage of Stickeres and Memes within a given Microsoft Teams team

```sh
graph teams funsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowStickersAndMemes true
```