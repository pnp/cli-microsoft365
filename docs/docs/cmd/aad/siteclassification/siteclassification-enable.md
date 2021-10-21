# aad siteclassification enable

Enables site classification configuration

## Usage

```sh
m365 aad siteclassification enable [options]
```

## Options

`-c, --classifications <classifications>`
: Comma-separated list of classifications to enable in the tenant

`-d, --defaultClassification <defaultClassification>`
: Classification to use by default

`-u, --usageGuidelinesUrl [usageGuidelinesUrl]`
: URL with usage guidelines for members

`-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]`
: URL with usage guidelines for guests

--8<-- "docs/cmd/_global.md"

## Examples

Enable site classification

```sh
m365 aad siteclassification enable --classifications "High, Medium, Low" --defaultClassification "Medium"
```

Enable site classification with a usage guidelines URL

```sh
m365 aad siteclassification enable --classifications "High, Medium, Low" --defaultClassification "Medium" --usageGuidelinesUrl "http://aka.ms/pnp"
```

Enable site classification with usage guidelines URLs for guests and members

```sh
m365 aad siteclassification enable --classifications "High, Medium, Low" --defaultClassification "Medium" --usageGuidelinesUrl "http://aka.ms/pnp" --guestUsageGuidelinesUrl "http://aka.ms/pnp"
```

## More information

- SharePoint "modern" sites classification: [https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification)