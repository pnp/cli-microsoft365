# pa solution reference add

Adds a project reference to the solution in the current directory

## Usage

```sh
m365 pa solution reference add [options]
```

## Options

`-h, --help`
: output usage information

`-p, --path <path>`
: The path to the referenced project

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

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
