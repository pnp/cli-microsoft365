# spo sitedesign task remove

Removes the specified site design scheduled for execution

## Usage

```sh
m365 spo sitedesign task remove [options]
```

## Options

`-h, --help`
: output usage information

`-i, --taskId <taskId>`
: The ID of the site design task to remove

`--confirm`
: Don't prompt for confirming removing the site design task

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Removes the specified site design task with taskId _6ec3ca5b-d04b-4381-b169-61378556d76e_ scheduled for execution without prompting confirmation

```sh
m365 spo sitedesign task remove --taskId 6ec3ca5b-d04b-4381-b169-61378556d76e --confirm
```

Removes the specified site design task with taskId _6ec3ca5b-d04b-4381-b169-61378556d76e_ scheduled for execution with prompt for confirmation before removing

```sh
m365 spo sitedesign task remove --taskId 6ec3ca5b-d04b-4381-b169-61378556d76e
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)