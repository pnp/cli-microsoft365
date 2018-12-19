# Upgrade project spfx-170-webpart-react to v1.7.1

Date: 2018-12-16

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.7.1. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i @microsoft/sp-core-library@1.7.1 -SE
```

File: [./package.json](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i @microsoft/sp-lodash-subset@1.7.1 -SE
```

File: [./package.json](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i @microsoft/sp-office-ui-fabric-core@1.7.1 -SE
```

File: [./package.json](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i @microsoft/sp-webpart-base@1.7.1 -SE
```

File: [./package.json](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i @microsoft/sp-build-web@1.7.1 -DE
```

File: [./package.json](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i @microsoft/sp-module-interfaces@1.7.1 -DE
```

File: [./package.json](./package.json)

### FN002003 @microsoft/sp-webpart-workbench | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench

Execute the following command:

```sh
npm i @microsoft/sp-webpart-workbench@1.7.1 -DE
```

File: [./package.json](./package.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.7.1"
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN014006 sourceMapPathOverrides in .vscode/launch.json | Recommended

In the .vscode/launch.json file, for each configuration, in the sourceMapPathOverrides property, add "webpack:///.././src/*": "${webRoot}/src/*"

In file [.vscode/launch.json](.vscode/launch.json) update the code as follows:

```json
{
  "configurations": [{
      "name": "Local workbench",
      "sourceMapPathOverrides": {
        "webpack:///.././src/*": "${webRoot}/src/*"
      }
    },
    {
      "name": "Hosted workbench",
      "sourceMapPathOverrides": {
        "webpack:///.././src/*": "${webRoot}/src/*"
      }
    }
  ]
}
```

File: [.vscode/launch.json](.vscode/launch.json)

### FN020001 @types/react | Required

Add resolution for package @types/react

In file [./package.json](./package.json) update the code as follows:

```json
{
  "resolutions": {
    "@types/react": "16.4.2"
  }
}
```

File: [./package.json](./package.json)

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
npm i @microsoft/sp-core-library@1.7.1 @microsoft/sp-lodash-subset@1.7.1 @microsoft/sp-office-ui-fabric-core@1.7.1 @microsoft/sp-webpart-base@1.7.1 -SE
npm i @microsoft/sp-build-web@1.7.1 @microsoft/sp-module-interfaces@1.7.1 @microsoft/sp-webpart-workbench@1.7.1 -DE
npm dedupe
```

### Modify files

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.7.1"
  }
}
```

#### [.vscode/launch.json](.vscode/launch.json)

In the .vscode/launch.json file, for each configuration, in the sourceMapPathOverrides property, add "webpack:///.././src/*": "${webRoot}/src/*":

```json
{
  "configurations": [{
      "name": "Local workbench",
      "sourceMapPathOverrides": {
        "webpack:///.././src/*": "${webRoot}/src/*"
      }
    },
    {
      "name": "Hosted workbench",
      "sourceMapPathOverrides": {
        "webpack:///.././src/*": "${webRoot}/src/*"
      }
    }
  ]
}
```

#### [./package.json](./package.json)

Add resolution for package @types/react:

```json
{
  "resolutions": {
    "@types/react": "16.4.2"
  }
}
```
