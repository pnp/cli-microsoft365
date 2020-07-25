# spfx project externalize

Externalizes SharePoint Framework project dependencies

## Usage

```sh
m365 spfx project externalize [options]
```

## Options

`-h, --help`
: output usage information

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text,md`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    Run this command in the folder where the project for which you want to externalize dependencies is located. This command doesn't change your project files.

## Remarks

!!! attention
    This command is in preview and could change once it's officially released. If you see any room for improvement, we'd love to hear from you at [https://github.com/pnp/cli-microsoft365/issues](https://github.com/pnp/cli-microsoft365/issues).

The `spfx project externalize` command helps you externalize your SharePoint Framework project dependencies using the [unpkg CDN](https://unpkg.com/).

This command doesn't change your project files. Instead, it gives you a report with all steps necessary to externalize your project dependencies. Externalizing project dependencies is error-prone, especially when it comes to updating your solution's code. This is why at this moment, this command produces a report that you can use yourself to perform the necessary changes and verify that everything is working as expected.

## Examples

Get instructions to externalize the current SharePoint Framework project dependencies and save the findings in a Markdown file

```sh
m365 spfx project externalize --output md > "deps-report.md"
```

Get instructions to externalize the current SharePoint Framework project dependencies and show the summary of the findings in the terminal

```sh
m365 spfx project externalize
```
