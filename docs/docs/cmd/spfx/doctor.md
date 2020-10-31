# spfx doctor

Verifies environment configuration for using the specific version of the SharePoint Framework

## Usage

```sh
m365 spfx doctor [options]
```

## Options

`-h, --help`
: output usage information

`-e, --env [env]`
: Version of SharePoint for which to check compatibility: `sp2016|sp2019|spo`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text,md`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    Checks ran by this command are based on what is officially supported by Microsoft. It's possible that using different package managers or packages versions will work just fine.

## Remarks

This commands helps you to verify if your environment meets all prerequisites for building solutions using a particular version of the SharePoint Framework.

The command starts by detecting the version of SharePoint Framework that you want to use. First, it looks at the current project. If you didn't run the command in the context of a SharePoint Framework project, the command will try to determine the SharePoint Framework version based on the SharePoint Framework Yeoman generator that you have installed either in the current directory or globally.

Based on the determined version of the SharePoint Framework, the command will look at other dependencies such as Node.js, npm, Yeoman, Gulp, React and TypeScript to verify if their meet the requirements of that particular version of the SharePoint Framework.

If you miss any required tools or use a version that doesn't meet the SharePoint Framework requirements, the command will give you a list of recommendation how to address these issues.

Next to verifying the readiness of your environment to use a particular version of the SharePoint Framework, you can also check if the version of the SharePoint Framework that you use is compatible with the specific version of SharePoint. Supported versions are `sp2016`, `sp2019` and `spo`.

## Examples

Verify if your environment meets the requirements to work with the SharePoint Framework

```sh
m365 spfx doctor
```

Verify if your environment meets the requirements to work with the SharePoint Framework and also if the version of the SharePoint Framework that you're using is compatible with SharePoint 2019

```sh
m365 spfx doctor --env sp2019
```
