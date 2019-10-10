# spfx project externalize

Externalizes SharePoint Framework project dependencies

## Usage

```sh
spfx project externalize [options]
```

## Options

Option|Description
------|-----------
`-f, --outputFile [outputFile]`|Path to the file where the upgrade report should be stored in
`-o, --output [output]`|Output type. `json|text|md`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Run this command in the folder where the project that you want to externalize is located. This command doesn't change your project files.

## Remarks

The `spfx project externalize` command helps you externalize your SharePoint Framework project dependencies using the unpkg CDN. 

This command doesn't change your project files. Instead, it gives you a report with all steps necessary to externalize your project dependencies. Externalizing project dependencies is error-prone, especially when it comes to updating your solution's code. This is why at this moment, this command produces a report that you can use yourself to perform the necessary changes and verify that everything is working as expected.

## Examples

Get instructions to externalize the current SharePoint Framework project dependencies and save the findings in a Markdown file

```sh
spfx project exteranlize --output md --outputFile externals.md
```

Get instructions to externalize the current SharePoint Framework project dependencies and show the summary of the findings in the shell

```sh
spfx project externalize
```