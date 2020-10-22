# pa solution reference add

Adds a project reference to the solution in the current directory

## Usage

```sh
m365 pa solution reference add [options]
```

## Options

`-p, --path <path>`
: The path to the referenced project

--8<-- "docs/cmd/_global.md"

## Remarks

This commands expects a CDS solution project in the current directory, and references a PowerApps component framework project.

The CDS solution project and the PowerApps component framework project cannot have the same name.

## Examples

Adds a reference inside the CDS Solution project in the current directory to the PowerApps component framework project at `./projects/ExampleProject`

```sh
m365 pa solution reference add --path ./projects/ExampleProject
```

## More information

- Create and build a custom component: [https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf)
