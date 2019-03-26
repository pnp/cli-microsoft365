# graph siteclassification set

Updates site classification configuration

## Usage

```sh
graph siteclassification set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-c, --classifications [classifications]`|Comma-separated list of classifications
`-d, --defaultClassification [defaultClassification]`|Classification to use by default
`-u, --usageGuidelinesUrl [usageGuidelinesUrl]`|URL with usage guidelines for members
`-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]`|URL with usage guidelines for guests
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Update Office 365 Tenant site classification configuration

```sh
graph siteclassification set --classifications "High, Medium, Low" --defaultClassification "Medium"
```

Update only the default classification

```sh
graph siteclassification set --defaultClassification "Low"
```

Update site classification with a usage guidelines URL

```sh
graph siteclassification set --usageGuidelinesUrl "http://aka.ms/pnp"
```

Update site classification with usage guidelines URLs for guests and members

```sh
graph siteclassification set --usageGuidelinesUrl "http://aka.ms/pnp" --guestUsageGuidelinesUrl "http://aka.ms/pnp"
```

## More information

- SharePoint "modern" sites classification: [https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification)