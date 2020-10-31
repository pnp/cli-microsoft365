# Upgrade project spfx-134-webpart-nolib to v1.4.0

Date: 2018-6-14

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.4.0.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm update @microsoft/sp-core-library@1.4.0
```

File: [./package.json](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm update @microsoft/sp-lodash-subset@1.4.0
```

File: [./package.json](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm update @microsoft/sp-office-ui-fabric-core@1.4.0
```

File: [./package.json](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm update @microsoft/sp-webpart-base@1.4.0
```

File: [./package.json](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm update @microsoft/sp-build-web@1.4.0
```

File: [./package.json](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm update @microsoft/sp-module-interfaces@1.4.0
```

File: [./package.json](./package.json)

### FN002003 @microsoft/sp-webpart-workbench | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench

Execute the following command:

```sh
npm update @microsoft/sp-webpart-workbench@1.4.0
```

File: [./package.json](./package.json)

### FN006002 package-solution.json includeClientSideAssets | Required

Update package-solution.json includeClientSideAssets

In file [./config/package-solution.json](./config/package-solution.json) update the code as follows:

```json
{
  "solution": {
    "includeClientSideAssets": true
  }
}
```

File: [./config/package-solution.json](./config/package-solution.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.4.0"
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN012003 tsconfig.json skipLibCheck | Required

Update skipLibCheck in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "skipLibCheck": true
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN012004 tsconfig.json typeRoots ./node_modules/@types | Required

Add ./node_modules/@types typeRoots in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@types"
    ]
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN012005 tsconfig.json typeRoots ./node_modules/@microsoft | Required

Add ./node_modules/@microsoft typeRoots in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@microsoft"
    ]
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN012007 tsconfig.json es5 lib | Required

Add es5 lib in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "lib": [
      "es5"
    ]
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN012008 tsconfig.json dom lib | Required

Add dom lib in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "lib": [
      "dom"
    ]
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN012009 tsconfig.json es2015.collection lib | Required

Add es2015.collection lib in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "lib": [
      "es2015.collection"
    ]
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN013001 gulpfile.js ms-Grid sass suppression | Recommended

Add suppression for ms-Grid sass warning

In file [./gulpfile.js](./gulpfile.js) update the code as follows:

```js
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
```

File: [./gulpfile.js](./gulpfile.js)

### FN014001 json.schemas in .vscode/settings.json | Recommended

Remove json.schemas in .vscode/settings.json

In file [.vscode/settings.json](.vscode/settings.json) update the code as follows:

```json

```

File: [.vscode/settings.json](.vscode/settings.json)

## Summary

### Execute script

```sh
npm update @microsoft/sp-core-library@1.4.0
npm update @microsoft/sp-lodash-subset@1.4.0
npm update @microsoft/sp-office-ui-fabric-core@1.4.0
npm update @microsoft/sp-webpart-base@1.4.0
npm update @microsoft/sp-build-web@1.4.0
npm update @microsoft/sp-module-interfaces@1.4.0
npm update @microsoft/sp-webpart-workbench@1.4.0
```

### Modify files

#### [./config/package-solution.json](./config/package-solution.json)

```json
{
  "solution": {
    "includeClientSideAssets": true
  }
}
```

#### [./.yo-rc.json](./.yo-rc.json)

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.4.0"
  }
}
```

#### [./tsconfig.json](./tsconfig.json)

```json
{
  "compilerOptions": {
    "skipLibCheck": true
  }
}
```

```json
{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@types"
    ]
  }
}
```

```json
{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@microsoft"
    ]
  }
}
```

```json
{
  "compilerOptions": {
    "lib": [
      "es5"
    ]
  }
}
```

```json
{
  "compilerOptions": {
    "lib": [
      "dom"
    ]
  }
}
```

```json
{
  "compilerOptions": {
    "lib": [
      "es2015.collection"
    ]
  }
}
```

#### [./gulpfile.js](./gulpfile.js)

```js
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
```

#### [.vscode/settings.json](.vscode/settings.json)

```json

```
