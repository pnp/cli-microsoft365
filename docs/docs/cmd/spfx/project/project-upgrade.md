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

The `spfx project upgrade` command helps you upgrade your SharePoint Framework project to the specified version. If no version is specified, the command will upgrade to the latest version of the SharePoint Framework it supports (v1.16.0-rc.0).

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
        "description": "Upgrade SharePoint Framework dependency package @microsoft/sp-property-pane",
        "id": "FN001021",
        "file": "./package.json",
        "position": {
          "line": 17,
          "character": 5
        },
        "resolution": "npm i -SE @microsoft/sp-property-pane@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-property-pane"
      },
      {
        "description": "Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base",
        "id": "FN001004",
        "file": "./package.json",
        "position": {
          "line": 18,
          "character": 5
        },
        "resolution": "npm i -SE @microsoft/sp-webpart-base@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-webpart-base"
      },
      {
        "description": "Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset",
        "id": "FN001002",
        "file": "./package.json",
        "position": {
          "line": 16,
          "character": 5
        },
        "resolution": "npm i -SE @microsoft/sp-lodash-subset@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-lodash-subset"
      },
      {
        "description": "Install SharePoint Framework dependency package @microsoft/sp-adaptive-card-extension-base",
        "id": "FN001034",
        "file": "./package.json",
        "position": {
          "line": 14,
          "character": 3
        },
        "resolution": "npm i -SE @microsoft/sp-adaptive-card-extension-base@1.15.2",
        "resolutionType": "cmd",
        "severity": "Optional",
        "title": "@microsoft/sp-adaptive-card-extension-base"
      },
      {
        "description": "Install SharePoint Framework dev dependency package @microsoft/eslint-plugin-spfx",
        "id": "FN002022",
        "file": "./package.json",
        "position": {
          "line": 25,
          "character": 3
        },
        "resolution": "npm i -DE @microsoft/eslint-plugin-spfx@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/eslint-plugin-spfx"
      },
      {
        "description": "Install SharePoint Framework dev dependency package @microsoft/eslint-config-spfx",
        "id": "FN002023",
        "file": "./package.json",
        "position": {
          "line": 25,
          "character": 3
        },
        "resolution": "npm i -DE @microsoft/eslint-config-spfx@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/eslint-config-spfx"
      },
      {
        "description": "Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web",
        "id": "FN002001",
        "file": "./package.json",
        "position": {
          "line": 27,
          "character": 5
        },
        "resolution": "npm i -DE @microsoft/sp-build-web@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-build-web"
      },
      {
        "description": "Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces",
        "id": "FN002002",
        "file": "./package.json",
        "position": {
          "line": 28,
          "character": 5
        },
        "resolution": "npm i -DE @microsoft/sp-module-interfaces@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-module-interfaces"
      },
      {
        "description": "Install SharePoint Framework dev dependency package typescript",
        "id": "FN002026",
        "file": "./package.json",
        "position": {
          "line": 25,
          "character": 3
        },
        "resolution": "npm i -DE typescript@4.5.5",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "typescript"
      },
      {
        "description": "Update version in .yo-rc.json",
        "id": "FN010001",
        "file": "./.yo-rc.json",
        "position": {
          "line": 7,
          "character": 5
        },
        "resolution": "{\\\n  \"@microsoft/generator-sharepoint\": {\\\n    \"version\": \"1.15.2\"\\\n  }\\\n}",
        "resolutionType": "json",
        "severity": "Recommended",
        "title": ".yo-rc.json version"
      },
      {
        "description": "Add noImplicitAny in tsconfig.json",
        "id": "FN012020",
        "file": "./tsconfig.json",
        "position": {
          "line": 3,
          "character": 22
        },
        "resolution": "{\\\n  \"compilerOptions\": {\\\n    \"noImplicitAny\": true\\\n  }\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "tsconfig.json noImplicitAny"
      },
      {
        "description": "Update serve.json schema URL",
        "id": "FN007001",
        "file": "./config/serve.json",
        "position": {
          "line": 2,
          "character": 3
        },
        "resolution": "{\\\n  \"$schema\": \"https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json\"\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "serve.json schema"
      },
      {
        "description": "Upgrade SharePoint Framework dependency package office-ui-fabric-react",
        "id": "FN001022",
        "file": "./package.json",
        "position": {
          "line": 21,
          "character": 5
        },
        "resolution": "npm i -SE office-ui-fabric-react@7.185.7",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "office-ui-fabric-react"
      },
      {
        "description": "Install SharePoint Framework dependency package tslib",
        "id": "FN001033",
        "file": "./package.json",
        "position": {
          "line": 14,
          "character": 3
        },
        "resolution": "npm i -SE tslib@2.3.1",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "tslib"
      },
      {
        "description": "Upgrade SharePoint Framework dev dependency package ajv",
        "id": "FN002007",
        "file": "./package.json",
        "position": {
          "line": 37,
          "character": 5
        },
        "resolution": "npm i -DE ajv@6.12.5",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "ajv"
      },
      {
        "description": "Remove SharePoint Framework dev dependency package @microsoft/sp-tslint-rules",
        "id": "FN002009",
        "file": "./package.json",
        "position": {
          "line": 29,
          "character": 5
        },
        "resolution": "npm un -D @microsoft/sp-tslint-rules",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-tslint-rules"
      },
      {
        "description": "Upgrade SharePoint Framework dev dependency package @types/webpack-env",
        "id": "FN002013",
        "file": "./package.json",
        "position": {
          "line": 36,
          "character": 5
        },
        "resolution": "npm i -DE @types/webpack-env@1.15.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@types/webpack-env"
      },
      {
        "description": "Install SharePoint Framework dev dependency package @microsoft/rush-stack-compiler-4.5",
        "id": "FN002020",
        "file": "./package.json",
        "position": {
          "line": 25,
          "character": 3
        },
        "resolution": "npm i -DE @microsoft/rush-stack-compiler-4.5@0.2.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/rush-stack-compiler-4.5"
      },
      {
        "description": "Install SharePoint Framework dev dependency package @rushstack/eslint-config",
        "id": "FN002021",
        "file": "./package.json",
        "position": {
          "line": 25,
          "character": 3
        },
        "resolution": "npm i -DE @rushstack/eslint-config@2.5.1",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@rushstack/eslint-config"
      },
      {
        "description": "Install SharePoint Framework dev dependency package eslint",
        "id": "FN002024",
        "file": "./package.json",
        "position": {
          "line": 25,
          "character": 3
        },
        "resolution": "npm i -DE eslint@8.7.0",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "eslint"
      },
      {
        "description": "Install SharePoint Framework dev dependency package eslint-plugin-react-hooks",
        "id": "FN002025",
        "file": "./package.json",
        "position": {
          "line": 25,
          "character": 3
        },
        "resolution": "npm i -DE eslint-plugin-react-hooks@4.3.0",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "eslint-plugin-react-hooks"
      },
      {
        "description": "Update tsconfig.json extends property",
        "id": "FN012017",
        "file": "./tsconfig.json",
        "position": {
          "line": 2,
          "character": 3
        },
        "resolution": "{\\\n  \"extends\": \"./node_modules/@microsoft/rush-stack-compiler-4.5/includes/tsconfig-web.json\"\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "tsconfig.json extends property"
      },
      {
        "description": "Remove file tslint.json",
        "id": "FN015003",
        "file": "tslint.json",
        "resolution": "rm \"tslint.json\"",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "tslint.json"
      },
      {
        "description": "Add file .eslintrc.js",
        "id": "FN015008",
        "file": ".eslintrc.js",
        "resolution": "cat > \".eslintrc.js\" << EOF \\\nrequire('@rushstack/eslint-config/patch/modern-module-resolution');\\\nmodule.exports = {\\\n  extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],\\\n  parserOptions: { tsconfigRootDir: __dirname }\\\n};\\\nEOF",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": ".eslintrc.js"
      },
      {
        "description": "To .gitignore add the '.heft' folder",
        "id": "FN023002",
        "file": "./.gitignore",
        "resolution": ".heft",
        "resolutionType": "text",
        "severity": "Required",
        "title": ".gitignore '.heft' folder"
      },
      {
        "description": "In package-solution.json add metadata section",
        "id": "FN006005",
        "file": "./config/package-solution.json",
        "position": {
          "line": 3,
          "character": 3
        },
        "resolution": "{\\\n  \"solution\": {\\\n    \"metadata\": {\\\n      \"shortDescription\": {\\\n        \"default\": \"react-youtube description\"\\\n      },\\\n      \"longDescription\": {\\\n        \"default\": \"react-youtube description\"\\\n      },\\\n      \"screenshotPaths\": [],\\\n      \"videoUrl\": \"\",\\\n      \"categories\": []\\\n    }\\\n  }\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "package-solution.json metadata"
      },
      {
        "description": "Upgrade SharePoint Framework dependency package react",
        "id": "FN001008",
        "file": "./package.json",
        "position": {
          "line": 22,
          "character": 5
        },
        "resolution": "npm i -SE react@16.13.1",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "react"
      },
      {
        "description": "Upgrade SharePoint Framework dependency package react-dom",
        "id": "FN001009",
        "file": "./package.json",
        "position": {
          "line": 23,
          "character": 5
        },
        "resolution": "npm i -SE react-dom@16.13.1",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "react-dom"
      },
      {
        "description": "Remove SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench",
        "id": "FN002003",
        "file": "./package.json",
        "position": {
          "line": 30,
          "character": 5
        },
        "resolution": "npm un -D @microsoft/sp-webpart-workbench",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@microsoft/sp-webpart-workbench"
      },
      {
        "description": "Upgrade SharePoint Framework dev dependency package @types/react",
        "id": "FN002015",
        "file": "./package.json",
        "position": {
          "line": 34,
          "character": 5
        },
        "resolution": "npm i -DE @types/react@16.9.51",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@types/react"
      },
      {
        "description": "Update serve.json initialPage URL",
        "id": "FN007002",
        "file": "./config/serve.json",
        "position": {
          "line": 5,
          "character": 3
        },
        "resolution": "{\\\n  \"initialPage\": \"https://enter-your-SharePoint-site/_layouts/workbench.aspx\"\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "serve.json initialPage"
      },
      {
        "description": "From serve.json remove the api property",
        "id": "FN007003",
        "file": "./config/serve.json",
        "position": {
          "line": 6,
          "character": 3
        },
        "resolution": "",
        "resolutionType": "json",
        "severity": "Required",
        "title": "serve.json api"
      },
      {
        "description": "Remove file config\\copy-assets.json",
        "id": "FN015007",
        "file": "config\\copy-assets.json",
        "resolution": "rm \"config\\copy-assets.json\"",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "config\\copy-assets.json"
      },
      {
        "description": "Create the .npmignore file",
        "id": "FN024001",
        "file": "./.npmignore",
        "resolution": "!dist\\\nconfig\\\n\ngulpfile.js\\\n\nrelease\\\nsrc\\\ntemp\\\n\ntsconfig.json\\\ntslint.json\\\n\n*.log\\\n\n.yo-rc.json\\\n.vscode\\\n",
        "resolutionType": "text",
        "severity": "Required",
        "title": "Create .npmignore"
      },
      {
        "description": "Update deploy-azure-storage.json workingDir",
        "id": "FN005002",
        "file": "./config/deploy-azure-storage.json",
        "position": {
          "line": 3,
          "character": 3
        },
        "resolution": "{\\\n  \"workingDir\": \"./release/assets/\"\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "deploy-azure-storage.json workingDir"
      },
      {
        "description": "To .gitignore add the 'release' folder",
        "id": "FN023001",
        "file": "./.gitignore",
        "resolution": "release",
        "resolutionType": "text",
        "severity": "Required",
        "title": ".gitignore 'release' folder"
      },
      {
        "description": "Upgrade SharePoint Framework dev dependency package gulp",
        "id": "FN002004",
        "file": "./package.json",
        "position": {
          "line": 38,
          "character": 5
        },
        "resolution": "npm i -DE gulp@4.0.2",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "gulp"
      },
      {
        "description": "Remove SharePoint Framework dev dependency package @types/chai",
        "id": "FN002005",
        "file": "./package.json",
        "position": {
          "line": 31,
          "character": 5
        },
        "resolution": "npm un -D @types/chai",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@types/chai"
      },
      {
        "description": "Remove SharePoint Framework dev dependency package @types/mocha",
        "id": "FN002006",
        "file": "./package.json",
        "position": {
          "line": 33,
          "character": 5
        },
        "resolution": "npm un -D @types/mocha",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@types/mocha"
      },
      {
        "description": "Remove SharePoint Framework dev dependency package @types/es6-promise",
        "id": "FN002014",
        "file": "./package.json",
        "position": {
          "line": 32,
          "character": 5
        },
        "resolution": "npm un -D @types/es6-promise",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@types/es6-promise"
      },
      {
        "description": "Upgrade SharePoint Framework dev dependency package @types/react-dom",
        "id": "FN002016",
        "file": "./package.json",
        "position": {
          "line": 35,
          "character": 5
        },
        "resolution": "npm i -DE @types/react-dom@16.9.8",
        "resolutionType": "cmd",
        "severity": "Required",
        "title": "@types/react-dom"
      },
      {
        "description": "Remove tsconfig.json exclude property",
        "id": "FN012013",
        "file": "./tsconfig.json",
        "position": {
          "line": 35,
          "character": 3
        },
        "resolution": "{\\\n  \"exclude\": []\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "tsconfig.json exclude property"
      },
      {
        "description": "Add es2015.promise lib in tsconfig.json",
        "id": "FN012018",
        "file": "./tsconfig.json",
        "position": {
          "line": 25,
          "character": 5
        },
        "resolution": "{\\\n  \"compilerOptions\": {\\\n    \"lib\": [\\\n      \"es2015.promise\"\\\n    ]\\\n  }\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "tsconfig.json es2015.promise lib"
      },
      {
        "description": "Remove es6-promise type in tsconfig.json",
        "id": "FN012019",
        "file": "./tsconfig.json",
        "position": {
          "line": 22,
          "character": 7
        },
        "resolution": "{\\\n  \"compilerOptions\": {\\\n    \"types\": [\\\n      \"es6-promise\"\\\n    ]\\\n  }\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "tsconfig.json es6-promise types"
      },
      {
        "description": "Before 'build.initialize(require('gulp'));' add the serve task",
        "id": "FN013002",
        "file": "./gulpfile.js",
        "resolution": "var getTasks = build.rig.getTasks;\\\nbuild.rig.getTasks = function () {\\\n  var result = getTasks.call(build.rig);\\\n\n  result.set('serve', result.get('serve-deprecated'));\\\n\n  return result;\\\n};\\\n",
        "resolutionType": "js",
        "severity": "Required",
        "title": "gulpfile.js serve task"
      },
      {
        "description": "Update tslint.json extends property",
        "id": "FN019002",
        "file": "./tslint.json",
        "position": {
          "line": 2,
          "character": 5
        },
        "resolution": "{\\\n  \"extends\": \"./node_modules/@microsoft/sp-tslint-rules/base-tslint.json\"\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "tslint.json extends"
      },
      {
        "description": "Remove package.json property",
        "id": "FN021002",
        "file": "./package.json",
        "position": {
          "line": 6,
          "character": 3
        },
        "resolution": "{\\\n  \"engines\": \"undefined\"\\\n}",
        "resolutionType": "json",
        "severity": "Required",
        "title": "engines"
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
    Execute in bash
    -----------------------
    npm un -D @microsoft/sp-tslint-rules @microsoft/sp-webpart-workbench @types/chai @types/mocha @types/es6-promise
    npm i -SE @microsoft/sp-core-library@1.15.2 @microsoft/sp-property-pane@1.15.2 @microsoft/sp-webpart-base@1.15.2 @microsoft/sp-lodash-subset@1.15.2 @microsoft/sp-adaptive-card-extension-base@1.15.2 office-ui-fabric-react@7.185.7 tslib@2.3.1 react@16.13.1 react-dom@16.13.1
    npm i -DE @microsoft/eslint-plugin-spfx@1.15.2 @microsoft/eslint-config-spfx@1.15.2 @microsoft/sp-build-web@1.15.2 @microsoft/sp-module-interfaces@1.15.2 typescript@4.5.5 ajv@6.12.5 @types/webpack-env@1.15.2 @microsoft/rush-stack-compiler-4.5@0.2.2 @rushstack/eslint-config@2.5.1 eslint@8.7.0 eslint-plugin-react-hooks@4.3.0 @types/react@16.9.51 gulp@4.0.2 @types/react-dom@16.9.8
    npm dedupe
    rm "tslint.json"
    cat > ".eslintrc.js" << EOF
    require('@rushstack/eslint-config/patch/modern-module-resolution');
    module.exports = {
      extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
      parserOptions: { tsconfigRootDir: __dirname }
    };
    EOF
    rm "config\copy-assets.json"

    ./.yo-rc.json
    -------------
    Update version in .yo-rc.json:
    {
      "@microsoft/generator-sharepoint": {
        "version": "1.15.2"
      }
    }


    ./tsconfig.json
    ---------------
    Add noImplicitAny in tsconfig.json:
    {
      "compilerOptions": {
        "noImplicitAny": true
      }
    }

    Update tsconfig.json extends property:
    {
      "extends": "./node_modules/@microsoft/rush-stack-compiler-4.5/includes/tsconfig-web.json"
    }

    Remove tsconfig.json exclude property:
    {
      "exclude": []
    }

    Add es2015.promise lib in tsconfig.json:
    {
      "compilerOptions": {
        "lib": [
          "es2015.promise"
        ]
      }
    }

    Remove es6-promise type in tsconfig.json:
    {
      "compilerOptions": {
        "types": [
          "es6-promise"
        ]
      }
    }


    ./config/serve.json
    -------------------
    Update serve.json schema URL:
    {
      "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json"
    }

    Update serve.json initialPage URL:
    {
      "initialPage": "https://enter-your-SharePoint-site/_layouts/workbench.aspx"
    }

    From serve.json remove the api property:



    ./.gitignore
    ------------
    To .gitignore add the '.heft' folder:
    .heft

    To .gitignore add the 'release' folder:
    release


    ./config/package-solution.json
    ------------------------------
    In package-solution.json add metadata section:
    {
      "solution": {
        "metadata": {
          "shortDescription": {
            "default": "react-youtube description"
          },
          "longDescription": {
            "default": "react-youtube description"
          },
          "screenshotPaths": [],
          "videoUrl": "",
          "categories": []
        }
      }
    }


    ./.npmignore
    ------------
    Create the .npmignore file:
    !dist
    config

    gulpfile.js

    release
    src
    temp

    tsconfig.json
    tslint.json

    *.log

    .yo-rc.json
    .vscode



    ./config/deploy-azure-storage.json
    ----------------------------------
    Update deploy-azure-storage.json workingDir:
    {
      "workingDir": "./release/assets/"
    }


    ./gulpfile.js
    -------------
    Before 'build.initialize(require('gulp'));' add the serve task:
    var getTasks = build.rig.getTasks;
    build.rig.getTasks = function () {
      var result = getTasks.call(build.rig);

      result.set('serve', result.get('serve-deprecated'));

      return result;
    };



    ./tslint.json
    -------------
    Update tslint.json extends property:
    {
      "extends": "./node_modules/@microsoft/sp-tslint-rules/base-tslint.json"
    }


    ./package.json
    --------------
    Remove package.json property:
    {
      "engines": "undefined"
    }
	  ```
