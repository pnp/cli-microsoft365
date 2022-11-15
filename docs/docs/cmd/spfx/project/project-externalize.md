# spfx project externalize

Externalizes SharePoint Framework project dependencies

## Usage

```sh
m365 spfx project externalize [options]
```

## Options

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

`-h, --help`
: output usage information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text,csv,md`. Default `json`

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

Supported SharePoint Framework versions are 1.0.0, 1.0.1, 1.0.2, 1.1.0, 1.1.1, 1.1.3, 1.2.0, 1.3.0, 1.3.1, 1.3.2, 1.3.4, 1.4.0, 1.4.1, 1.5.0, 1.5.1, 1.6.0, 1.7.0, 1.7.1, 1.8.0, 1.8.1, 1.8.2, 1.9.1.

## Examples

Get instructions to externalize the current SharePoint Framework project dependencies and save the findings in a Markdown file

```sh
m365 spfx project externalize --output md > "deps-report.md"
```

Get instructions to externalize the current SharePoint Framework project dependencies and show the summary of the findings in the terminal

```sh
m365 spfx project externalize
```

## Response

Below output will be produced to externalize the SharePoint Framework project dependencies.

=== "JSON"

    ```json
    {
      "externalConfiguration": {
        "externals": {
          "@pnp/odata": {
            "path": "https://unpkg.com/@pnp/odata@^1.3.11/dist/odata.es5.umd.min.js",
            "globalName": "pnp.odata",
            "globalDependencies": [
              "@pnp/common",
              "@pnp/logging",
              "tslib"
            ]
          },
          "@pnp/common": {
            "path": "https://unpkg.com/@pnp/common@^1.3.11/dist/common.es5.umd.bundle.min.js",
            "globalName": "pnp.common"
          },
          "@pnp/logging": {
            "path": "https://unpkg.com/@pnp/logging@^1.3.11/dist/logging.es5.umd.min.js",
            "globalName": "pnp.logging",
            "globalDependencies": [
              "tslib"
            ]
          },
          "@pnp/sp": {
            "path": "https://unpkg.com/@pnp/sp@^1.3.11/dist/sp.es5.umd.min.js",
            "globalName": "pnp.sp",
            "globalDependencies": [
              "@pnp/logging",
              "@pnp/common",
              "@pnp/odata",
              "tslib"
            ]
          },
          "tslib": {
            "path": "https://unpkg.com/tslib@^1.10.0/tslib.js",
            "globalName": "tslib"
          }
        }
      },
      "edits": [
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"@pnp/odata\");"
        },
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"@pnp/common\");"
        },
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"@pnp/logging\");"
        },
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"tslib\");"
        }
      ]
    }
    ```


=== "Text"

    ```text
    In the config/config.json file update the externals property to:

    {
      "externalConfiguration": {
        "externals": {
          "@pnp/odata": {
            "path": "https://unpkg.com/@pnp/odata@^1.3.11/dist/odata.es5.umd.min.js",
            "globalName": "pnp.odata",
            "globalDependencies": [
              "@pnp/common",
              "@pnp/logging",
              "tslib"
            ]
          },
          "@pnp/common": {
            "path": "https://unpkg.com/@pnp/common@^1.3.11/dist/common.es5.umd.bundle.min.js",
            "globalName": "pnp.common"
          },
          "@pnp/logging": {
            "path": "https://unpkg.com/@pnp/logging@^1.3.11/dist/logging.es5.umd.min.js",
            "globalName": "pnp.logging",
            "globalDependencies": [
              "tslib"
            ]
          },
          "@pnp/sp": {
            "path": "https://unpkg.com/@pnp/sp@^1.3.11/dist/sp.es5.umd.min.js",
            "globalName": "pnp.sp",
            "globalDependencies": [
              "@pnp/logging",
              "@pnp/common",
              "@pnp/odata",
              "tslib"
            ]
          },
          "tslib": {
            "path": "https://unpkg.com/tslib@^1.10.0/tslib.js",
            "globalName": "tslib"
          }
        }
      },
      "edits": [
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"@pnp/odata\");"
        },
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"@pnp/common\");"
        },
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"@pnp/logging\");"
        },
        {
          "action": "add",
          "path": "C:\\react-global-news-sp2019\\src\\webparts\\news\\NewsWebPart.ts",
          "targetValue": "require(\"tslib\");"
        }
      ]
    }
    ```

=== "Markdown"

    ````
    # Externalizing dependencies of project react-global-news-sp2019

    Date: 11/7/2022

    ## Findings

    ### Modify files

    #### [config.json](config/config.json)

    Replace the externals property (or add if not defined) with
    
    ```json
    {
      "externals": {
        "@pnp/odata": {
          "path": "https://unpkg.com/@pnp/odata@^1.3.11/dist/odata.es5.umd.min.js",
          "globalName": "pnp.odata",
          "globalDependencies": [
            "@pnp/common",
            "@pnp/logging",
            "tslib"
          ]
        },
        "@pnp/common": {
          "path": "https://unpkg.com/@pnp/common@^1.3.11/dist/common.es5.umd.bundle.min.js",
          "globalName": "pnp.common"
        },
        "@pnp/logging": {
          "path": "https://unpkg.com/@pnp/logging@^1.3.11/dist/logging.es5.umd.min.js",
          "globalName": "pnp.logging",
          "globalDependencies": [
            "tslib"
          ]
        },
        "@pnp/sp": {
          "path": "https://unpkg.com/@pnp/sp@^1.3.11/dist/sp.es5.umd.min.js",
          "globalName": "pnp.sp",
          "globalDependencies": [
            "@pnp/logging",
            "@pnp/common",
            "@pnp/odata",
            "tslib"
          ]
        },
        "tslib": {
          "path": "https://unpkg.com/tslib@^1.10.0/tslib.js",
          "globalName": "tslib"
        }
      }
    }
    ```
    
    #### [C:\react-global-news-sp2019\src\webparts\news\NewsWebPart.ts](C:\react-global-news-sp2019\src\webparts\news\NewsWebPart.ts)
    add
    ```JavaScript
    require("@pnp/odata");
    require("@pnp/common");
    require("@pnp/logging");
    require("tslib");
    ```
    ````
