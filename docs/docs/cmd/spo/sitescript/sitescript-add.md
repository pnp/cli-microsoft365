# spo sitescript add

Adds site script for use with site designs

## Usage

```sh
m365 spo sitescript add [options]
```

## Options

`-h, --help`
: output usage information

`-t, --title <title>`
: Site script title

`-c, --content <content>`
: JSON string containing the site script

`-d, --description [description]`
: Site script description

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Each time you execute the `spo sitescript add` command, it will create a new site script with a unique ID. Before creating a site script, be sure that another script with the same name doesn't already exist.

## Examples

Create new site script for use with site designs. Script contents are stored in the `$script` variable

```sh
m365 spo sitescript add --title "Contoso" --description "Contoso theme script" --content $script
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)