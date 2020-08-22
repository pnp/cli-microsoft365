# spfx project upgrade

Upgrades SharePoint Framework project to the specified version

## Usage

```sh
spfx project upgrade [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-v, --toVersion [toVersion]`|The version of SharePoint Framework to which upgrade the project
`--packageManager [packageManager]`|The package manager you use. Supported managers `npm,pnpm,yarn`. Default `npm`
`--shell [shell]`|The shell you use. Supported shells `bash,powershell,cmd`. Default `bash`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text,md,tour`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Run this command in the folder where the project that you want to upgrade is located. This command doesn't change your project files.

## Remarks

The `spfx project upgrade` command helps you upgrade your SharePoint Framework project to the specified version. If no version is specified, the command will upgrade to the latest version of the SharePoint Framework it supports (v1.11.0).

This command doesn't change your project files. Instead, it gives you a report with all steps necessary to upgrade your project to the specified version of the SharePoint Framework. Changing project files is error-prone, especially when it comes to updating your solution's code. This is why at this moment, this command produces a report that you can use yourself to perform the necessary updates and verify that everything is working as expected.

Using this command you can upgrade SharePoint Framework projects built using versions: 1.0.0, 1.0.1, 1.0.2, 1.1.0, 1.1.1, 1.1.3, 1.2.0, 1.3.0, 1.3.1, 1.3.2, 1.3.4, 1.4.0, 1.4.1, 1.5.0, 1.5.1, 1.6.0, 1.7.0, 1.7.1, 1.8.0, 1.8.1, 1.8.2, 1.9.1 and 1.10.0.

## Examples

Get instructions to upgrade the current SharePoint Framework project to SharePoint Framework version 1.5.0 and save the findings in a Markdown file

```sh
spfx project upgrade --toVersion 1.5.0 --output md > "upgrade-report.md"
```

Get instructions to Upgrade the current SharePoint Framework project to SharePoint Framework version 1.5.0 and show the summary of the findings in the shell

```sh
spfx project upgrade --toVersion 1.5.0
```

Get instructions to upgrade the current SharePoint Framework project to the latest SharePoint Framework version supported by the CLI for Microsoft 365 using pnpm

```sh
spfx project upgrade --packageManager pnpm
```

Get instructions to upgrade the current SharePoint Framework project to the latest SharePoint Framework version supported by the CLI for Microsoft 365

```sh
spfx project upgrade
```

Get instructions to upgrade the current SharePoint Framework project to the latest SharePoint Framework version supported by the CLI for Microsoft 365 using PowerShell

```sh
spfx project upgrade --shell powershell
```

Get instructions to upgrade the current SharePoint Framework project to the latest version of SharePoint Framework and save the findings in a [CodeTour](https://aka.ms/codetour) file

```sh
spfx project upgrade  --output tour
```
