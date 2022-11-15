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

## Response

### Response with no issues

=== "JSON"

    ```json
    []
    ```

=== "Text"

    ```text
    ✅ CLI for Microsoft 365 has found no issues in your project
    ```

=== "Markdown"

    ````
    # Validate project spfx-solution

    Date: 11/7/2022

    ## Findings

    ✅ CLI for Microsoft 365 has found no issues in your project
    ````

### Response with issues reported

When the npm packages related issues are reported. 

=== "JSON"

    ```json
    [
      {
        "description": "Package @microsoft/rush-stack-compiler-4.2 is installed as a dependency. Install it as a devDependency instead",
        "id": "FN021006",
        "file": "./package.json",
        "position": {
          "line": 14,
          "character": 19
        },
        "resolution": "npm i -DE @microsoft/rush-stack-compiler-4.2@^0.1.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/rush-stack-compiler-4.2 installed as a dependency"
      },
      {
        "description": "Install supported version of the office-ui-fabric-react package",
        "id": "FN001022",
        "file": "./package.json",
        "position": {
          "line": 24,
          "character": 5
        },
        "resolution": "npm i -SE office-ui-fabric-react@7.174.1",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "office-ui-fabric-react"
      },
      {
        "description": "Uninstall unsupported version of @microsoft/rush-stack-compiler",
        "id": "FN002019",
        "file": "./package.json",
        "position": {
          "line": 14,
          "character": 19
        },
        "resolution": "npm un -D @microsoft/rush-stack-compiler-4.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/rush-stack-compiler-3.9"
      },
      {
        "description": "If, after upgrading npm packages, when building the project you have errors similar to: \"error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'\", try running 'npm dedupe' to cleanup npm packages.",
        "id": "FN017001",
        "file": "./package.json",
        "resolution": "npm dedupe",
        "resolutionType": "cmd",
        "severity": "Optional",
        "title": "Run npm dedupe"
      }
    ]
    ```

=== "Text"

    ```text
    Execute in command line
    -----------------------
    npm un -D @microsoft/rush-stack-compiler-4.2
    npm i -SE office-ui-fabric-react@7.174.1
    npm i -DE @microsoft/rush-stack-compiler-4.2@^0.1.2
    npm dedupe
    ```

=== "Markdown"

    ````
    # Validate project react-page-navigator

    Date: 11/15/2022

    ## Findings

    Following is the list of issues found in your project. [Summary](#Summary) of the recommended fixes is included at the end of the report.

    ### FN021006 @microsoft/rush-stack-compiler-4.2 installed as a dependency | Required

    Package @microsoft/rush-stack-compiler-4.2 is installed as a dependency. Install it as a devDependency instead

    Execute the following command:

    ```sh
    npm i -DE @microsoft/rush-stack-compiler-4.2@^0.1.2
    ```

    File: [./package.json:14:19](./package.json)

    ### FN001022 office-ui-fabric-react | Required

    Install supported version of the office-ui-fabric-react package

    Execute the following command:

    ```sh
    npm i -SE office-ui-fabric-react@7.174.1
    ```

    File: [./package.json:24:5](./package.json)

    ### FN002019 @microsoft/rush-stack-compiler-3.9 | Required

    Uninstall unsupported version of @microsoft/rush-stack-compiler

    Execute the following command:

    ```sh
    npm un -D @microsoft/rush-stack-compiler-4.2
    ```

    File: [./package.json:14:19](./package.json)

    ### FN017001 Run npm dedupe | Optional

    If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

    Execute the following command:

    ```sh
    npm dedupe
    ```

    File: [./package.json](./package.json)

    ## Summary

    ### Execute script

    ```sh
    npm un -D @microsoft/rush-stack-compiler-4.2
    npm i -SE office-ui-fabric-react@7.174.1
    npm i -DE @microsoft/rush-stack-compiler-4.2@^0.1.2
    npm dedupe
    ```
    ````
