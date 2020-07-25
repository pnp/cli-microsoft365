# spo sitedesign task get

Gets information about the specified site design scheduled for execution

## Usage

```sh
m365 spo sitedesign task get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --taskId <taskId>`
: The ID of the site design task to get information for

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get information about the specified site design scheduled for execution

```sh
m365 spo sitedesign task get --taskId 6ec3ca5b-d04b-4381-b169-61378556d76e
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)