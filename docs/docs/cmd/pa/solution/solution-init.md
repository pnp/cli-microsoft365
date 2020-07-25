# pa solution init

Initializes a directory with a new CDS solution project

## Usage

```sh
m365 pa solution init [options]
```

## Options

`-h, --help`
: output usage information

`--publisherName <publisherName>`
: Name of the CDS solution publisher.

`--publisherPrefix <publisherPrefix>`
: Customization prefix value for the CDS solution publisher.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

PublisherName only allows characters within the ranges `[A-Z]`, `[a-z]`, `[0-9]`, or `_`. The first character may only be in the ranges `[A-Z]`, `[a-z]`, or `_`.

PublisherPrefix must be 2 to 8 characters long, can only consist of alpha-numerics, must start with a letter, and cannot start with 'mscrm'.

## Examples

Initializes a CDS solution project using _yourPublisherName_ as publisher name and _ypn_ as publisher prefix

```sh
m365 pa solution init --publisherName yourPublisherName --publisherPrefix ypn
```

## More information

- Create and build a custom component: [https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf](https://docs.microsoft.com/en-us/powerapps/developer/component-framework/create-custom-controls-using-pcf)
