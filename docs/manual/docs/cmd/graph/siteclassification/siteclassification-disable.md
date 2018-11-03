# graph siteclassification disable

Disables site classification

## Usage

```sh
graph siteclassification disable [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--confirm`|Don't prompt for confirming disabling site classification
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.
  
## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To disable site classification, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Disable site classification

```sh
graph siteclassification disable
```

Disable site classification without confirmation

```sh
graph siteclassification disable --confirm
```

## More information

- SharePoint "modern" sites classification: [https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification)