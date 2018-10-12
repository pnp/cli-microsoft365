# graph siteclassification enable

  Enables site classification configuration

## Usage 

```sh
graph siteclassification enable [options]
```

## Options 

Option|Description
------|-----------
    `--help`                                                   output usage information
    `-c, --classifications <classifications>`                  Comma-separated list of classifications to enable in the tenant
    `-d, --defaultClassification <defaultClassification>`      classification to use by default
    `-u, --usageGuidelinesUrl [usageGuidelinesUrl]`            URL with usage guidelines for members
    `-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]`  URL with usage guidelines for guests
    `-o, --output [output]`                                    Output type. `json|text.` Default `text`
    `--verbose`                                                Runs command with verbose logging
    `--debug`                                                  Runs command with debug logging

  !!!Important: before using this command, connect to the Microsoft Graph
    using the graph connect command.

## Remarks
    Attention: This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.

    To set the Office 365 Tenant site classification, you have
    to first connect to the Microsoft Graph using the graph connect command,
    eg. o365$ graph connect.

## Examples

    Enables SiteClassification
    ```sh
      o365$ graph siteclassification enable -c "High, Medium, Low" -d "Medium"
    ```
    
    Enables SiteClassification with a Usage Guidelines Url
    ```sh    
      o365$ graph siteclassification enable -c "High, Medium, Low" -d "Medium" --usageGuidelinesUrl "http://aka.ms/pnp"
    ```
    
    Enables SiteClassification with a Usage Guidelines Url and a Guestusage Guidelines Url
    ```sh    
      o365$ graph siteclassification enable -c "High, Medium, Low" -d "Medium" --usageGuidelinesUrl "http://aka.ms/pnp" --guestUsageGuidelinesUrl "http://aka.ms/pnp"
    ```

  ## More information

  - SharePoint "modern" sites classification
      [https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification]https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification