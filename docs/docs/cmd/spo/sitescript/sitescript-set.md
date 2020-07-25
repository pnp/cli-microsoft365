# spo sitescript set

Updates existing site script

## Usage

```sh
m365 spo sitescript set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: Site script ID

`-t, --title [title]`
: Site script title

`-d, --description [description]`
: Site script description

`-v, --version [version]`
: Site script version

`-c, --content [content]`
: JSON string containing the site script

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified `id` doesn't refer to an existing site script, you will get a `File not found` error.

## Examples

Update title of the existing site script with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
m365 spo sitescript set --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --title "Contoso"
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)