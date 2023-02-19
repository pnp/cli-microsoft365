# spfx doctor

Verifies environment configuration for using the specific version of the SharePoint Framework

## Usage

```sh
m365 spfx doctor [options]
```

## Options

`-e, --env [env]`
: Version of SharePoint for which to check compatibility: `sp2016|sp2019|spo`

`-v, --spfxVersion [spfxVersion]`
: Version of the SharePoint Framework Yeoman generator to check compatibility for without `v`, eg. `1.11.0`

`-h, --help`
: output usage information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    Checks ran by this command are based on what is officially supported by Microsoft. It's possible that using different package managers or packages versions will work just fine.

## Remarks

This commands helps you to verify if your environment meets all prerequisites for building solutions using a particular version of the SharePoint Framework.

The command starts by detecting the version of SharePoint Framework that you want to use. First, it looks at the current project. If you didn't run the command in the context of a SharePoint Framework project, the command will try to determine the SharePoint Framework version based on the SharePoint Framework Yeoman generator that you have installed either in the current directory or globally.

Based on the determined version of the SharePoint Framework, the command will look at other dependencies such as Node.js, npm, Yeoman, Gulp CLI and TypeScript to verify if their meet the requirements of that particular version of the SharePoint Framework.

If you miss any required tools or use a version that doesn't meet the SharePoint Framework requirements, the command will give you a list of recommendation how to address these issues.

Next to verifying the readiness of your environment to use a particular version of the SharePoint Framework, you can also check if the version of the SharePoint Framework that you use is compatible with the specific version of SharePoint. Supported versions are `sp2016`, `sp2019` and `spo`.

!!! important
    This command supports only text output.

## Examples

Verify if your environment meets the requirements to work with SharePoint Framework based on the globally installed version of the SharePoint Framework Yeoman generator or the current project

```sh
m365 spfx doctor --output text
```

Verify if your environment meets the requirements to work with the SharePoint Framework and also if the version of the SharePoint Framework that you're using is compatible with SharePoint 2019

```sh
m365 spfx doctor --env sp2019 --output text
```

Verify if your environment meets the requirements to work with SharePoint Framework v1.11.0

```sh
m365 spfx doctor --spfxVersion 1.11.0 --output text
```

## Response

### Response with no issues

=== "Text"

    ```text
    CLI for Microsoft 365 SharePoint Framework doctor
    Verifying configuration of your system for working with the SharePoint Framework

    √ SharePoint Framework v1.15.0
    √ Node v16.13.0    
    √ yo v4.3.0
    √ gulp-cli v2.3.0
    √ bundled typescript used
    ```

### Response with issues reported

When the installed version of Yeoman is lower than expected to run SharePoint Framework v1.15.0

=== "Text"

    ```text
    CLI for Microsoft 365 SharePoint Framework doctor
    Verifying configuration of your system for working with the SharePoint Framework

    √ SharePoint Framework v1.15.0
    √ Node v16.16.0
    × yo v3.1.1 found, v^4 required
    √ gulp-cli v2.3.0
    √ bundled typescript used

    Recommended fixes:

    - npm i -g yo@4
    ```
