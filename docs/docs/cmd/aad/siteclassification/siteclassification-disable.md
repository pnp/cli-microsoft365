# aad siteclassification disable

Disables site classification

## Usage

```sh
aad siteclassification disable [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--confirm`|Don't prompt for confirming disabling site classification
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Disable site classification

```sh
aad siteclassification disable
```

Disable site classification without confirmation

```sh
aad siteclassification disable --confirm
```

## More information

- SharePoint "modern" sites classification: [https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification)