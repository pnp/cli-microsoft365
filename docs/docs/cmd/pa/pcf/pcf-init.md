# pa pcf init

Creates new PowerApps component framework project

## Usage

```sh
m365 pa pcf init [options]
```

## Options

`-h, --help`
: output usage information

`--namespace <namespace>`
: The namespace for the component.

`--name <name>`
: The name for the component.

`--template <template>`
: Choose a template for the component. `Field,Dataset`.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Name cannot contain reserved Javascript words. Only characters within the ranges [A-Z], [a-z] or [0-9] are allowed. The first character may not be a number.

Namespace cannot contain reserved Javascript words. Only characters within the ranges [A-Z], [a-z], [0-9], or '.' are allowed. The first and last character may not be the '.' character. Consecutive '.' characters are not allowed. Numbers are not allowed as the first character or immediately after a period.

Template currently only supports Field or Dataset.

## Examples

Initialize the PowerApps Component Framework for a Field component

```sh
m365 pa pcf init --namespace yourNamespace --name yourCustomFieldComponent --template Field
```

Initialize the PowerApps Component Framework for a Dataset component

```sh
m365 pa pcf init --namespace yourNamespace --name yourCustomFieldComponent --template Dataset
```

## More information

- Create and build a custom component: [https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf)
