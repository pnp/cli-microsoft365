# aad siteclassification set

Updates site classification configuration

## Usage

```sh
m365 aad siteclassification set [options]
```

## Options

`-c, --classifications [classifications]`
: Comma-separated list of classifications

`-d, --defaultClassification [defaultClassification]`
: Classification to use by default

`-u, --usageGuidelinesUrl [usageGuidelinesUrl]`
: URL with usage guidelines for members

`-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]`
: URL with usage guidelines for guests

--8<-- "docs/cmd/_global.md"


## Examples

Update Microsoft 365 Tenant site classification configuration

```sh
m365 aad siteclassification set --classifications "High, Medium, Low" --defaultClassification "Medium"
```

Update only the default classification

```sh
m365 aad siteclassification set --defaultClassification "Low"
```

Update site classification with a usage guidelines URL

```sh
m365 aad siteclassification set --usageGuidelinesUrl "http://aka.ms/pnp"
```

Update site classification with usage guidelines URLs for guests and members

```sh
m365 aad siteclassification set --usageGuidelinesUrl "http://aka.ms/pnp" --guestUsageGuidelinesUrl "http://aka.ms/pnp"
```

## More information

- SharePoint "modern" sites classification: [https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification)
