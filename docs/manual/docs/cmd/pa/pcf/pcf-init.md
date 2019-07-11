# pa pcf init

Initializes a directory with a new PowerApps component framework project

## Usage

```sh
pa pcf init [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--name <name>`|The name for the component.
`--namespace <namespace>`|The namespace for the component.
`--template [template]`|Choose a template for the component. Field|Dataset.
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

More information in the Microsoft documentation about [creating and building a custom component for PowerApps](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf).

## Examples

Initialize the PowerApps Component Framework for a Field component

```sh
pa pcf init --namespace yourNamespace --name yourCustomFieldComponent --template Field
```

Initialize the PowerApps Component Framework for a Dataset component

```sh
pa pcf init --namespace yourNamespace --name yourCustomFieldComponent --template Dataset
```