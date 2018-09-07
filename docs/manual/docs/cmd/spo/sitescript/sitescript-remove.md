# spo sitescript remove

Removes the specified site script

## Usage

```sh
spo sitescript remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|Site script ID
`--confirm`|Don't prompt for confirming removing the site script
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To remove a site script, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the specified `id` doesn't refer to an existing site script, you will get a `File not found` error.

## Examples

Remove site script with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_. Will prompt for confirmation before removing the script

```sh
spo sitescript remove --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

Remove site script with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_ without prompting for confirmation

```sh
spo sitescript remove --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --confirm
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)