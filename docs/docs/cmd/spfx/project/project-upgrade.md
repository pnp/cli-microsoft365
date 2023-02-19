# spfx project upgrade

Upgrades SharePoint Framework project to the specified version

## Usage

```sh
m365 spfx project upgrade [options]
```

## Options

`-v, --toVersion [toVersion]`
: The version of SharePoint Framework to which upgrade the project

`--packageManager [packageManager]`
: The package manager you use. Supported managers `npm,pnpm,yarn`. Default `npm`

`--shell [shell]`
: The shell you use. Supported shells `bash,powershell,cmd`. Default `bash`

`--preview`
: Upgrade project to the latest SPFx preview version

`-f, --outputFile [outputFile]`
: Path to the file where the upgrade report should be stored in. Ignored when `output` is `tour`

--8<-- "docs/cmd/_global.md"

!!! important
    Run this command in the folder where the project that you want to upgrade is located. This command doesn't change your project files.

## Remarks

The `spfx project upgrade` command helps you upgrade your SharePoint Framework project to the specified version. If no version is specified, the command will upgrade to the latest version of the SharePoint Framework it supports (v1.16.1).

This command doesn't change your project files. Instead, it gives you a report with all steps necessary to upgrade your project to the specified version of the SharePoint Framework. Changing project files is error-prone, especially when it comes to updating your solution's code. This is why at this moment, this command produces a report that you can use yourself to perform the necessary updates and verify that everything is working as expected.

## Examples

Get instructions to upgrade the current SharePoint Framework project to SharePoint Framework version 1.5.0 and save the findings in a Markdown file

```sh
m365 spfx project upgrade --toVersion 1.5.0 --output md > "upgrade-report.md"
```

Get instructions to upgrade the current SharePoint Framework project to SharePoint Framework version 1.5.0 and show the summary of the findings in the shell

```sh
m365 spfx project upgrade --toVersion 1.5.0 --output text
```

Get instructions to upgrade the current SharePoint Framework project to the latest preview version

```sh
m365 spfx project upgrade --preview --output text
```

Get instructions to upgrade the current SharePoint Framework project to the specified preview version

```sh
m365 spfx project upgrade --toVersion 1.12.1-rc.0 --output text
```

Get instructions to upgrade the current SharePoint Framework project to the latest SharePoint Framework version supported by the CLI for Microsoft 365 using pnpm

```sh
m365 spfx project upgrade --packageManager pnpm --output text
```

Get instructions to upgrade the current SharePoint Framework project to the latest SharePoint Framework version supported by the CLI for Microsoft 365

```sh
m365 spfx project upgrade --output text
```

Get instructions to upgrade the current SharePoint Framework project to the latest SharePoint Framework version supported by the CLI for Microsoft 365 using PowerShell

```sh
m365 spfx project upgrade --shell powershell --output text
```

Get instructions to upgrade the current SharePoint Framework project to the latest version of SharePoint Framework and save the findings in a [CodeTour](https://aka.ms/codetour) file

```sh
m365 spfx project upgrade --output tour
```

## Response

When upgrading an SPFx project built using version 1.15.0 to SPFx version 1.15.2, you'll get output similar to following (output is truncated):

=== "JSON"

    ```json
    [
      {
        "description": "Upgrade SharePoint Framework dependency package @microsoft/sp-core-library",
        "id": "FN001001",
        "file": "./package.json",
        "position": {
          "line": 15,
          "character": 5
        },
        "resolution": "npm i -SE @microsoft/sp-core-library@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-core-library"
      },
      {
        "description": "Update version in .yo-rc.json",
        "id": "FN010001",
        "file": "./.yo-rc.json",
        "position": {
          "line": 5,
          "character": 5
        },
        "resolution": "{\\\n  \"@microsoft/generator-sharepoint\": {\\\n    \"version\": \"1.15.2\"\\\n  }\\\n}",
        "resolutionType": "json",
        "severity": "Recommended",
        "title": ".yo-rc.json version"
      }
    ]
    ```

=== "Text"

    ```text
    Execute in bash
    -----------------------
    npm i -SE @microsoft/sp-core-library@1.15.2

    ./.yo-rc.json
    -------------
    Update version in .yo-rc.json:
    {
      "@microsoft/generator-sharepoint": {
        "version": "1.15.2"
      }
    }
    ```

=== "Markdown"

    ````md
    # Upgrade project HelloWorld to v1.15.2

    Date: 20/11/2022

    ## Findings

    Following is the list of steps required to upgrade your project to SharePoint Framework version 1.15.2. [Summary](#Summary) of the modifications is included at the end of the report.

    ### FN001001 @microsoft/sp-core-library | Required

    Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

    Execute the following command:

    ```sh
    npm i -SE @microsoft/sp-core-library@1.15.2
    ```

    File: [./package.json:17:5](./package.json)

    ## Summary

    ### Execute script

    ```sh
    npm i -SE @microsoft/sp-core-library@1.15.2
    ```

    ### Modify files

    #### [./.yo-rc.json](./.yo-rc.json)

    Update version in .yo-rc.json:

    ```json
    {
      "@microsoft/generator-sharepoint": {
        "version": "1.15.2"
      }
    }
    ```
    ````
