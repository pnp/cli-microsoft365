# spfx project doctor

Validates correctness of a SharePoint Framework project

## Usage

```sh
m365 spfx project doctor [options]
```

## Options

`--packageManager [packageManager]`
: The package manager you use. Supported managers `npm,pnpm,yarn`. Default `npm`

`-h, --help`
: output usage information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text,tour,csv`. Default `json`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    Run this command in the folder where the project that you want to validate is located. This command doesn't change your project files.

## Remarks

The `spfx project doctor` command helps you validate that your SharePoint Framework project is set up correctly. The command automatically detects the version of your project using version information specified in the project's .yo-rc.json file or package.json (if no version information is included in .yo-rc.json). Based on the detected project version, the command executes several checks and reports any issues in the specified format.

This command doesn't change your project files. Instead, it gives you a report with all steps necessary to validate your project to the specified version of the SharePoint Framework. Changing project files is error-prone, especially when it comes to updating your solution's code. This is why at this moment, this command produces a report that you can use yourself to perform the necessary updates and verify that everything is working as expected.

## Examples

Validate if your project is correctly set up and save the findings in a Markdown file

```sh
m365 spfx project doctor --output md > "doctor-report.md"
```

Validate if your project is correctly set up and show the summary of the findings in the terminal

```sh
m365 spfx project doctor --output text
```

Validate if your project is correctly set up and get instructions to fix any issues using pnpm

```sh
m365 spfx project doctor --packageManager pnpm --output text
```

Validate if your project is correctly set up and get instructions to fix any issues in a [CodeTour](https://aka.ms/codetour) file

```sh
m365 spfx project doctor --output tour
```
