# spo sitedesign apply

Applies a site design to an existing site collection

## Usage

```sh
m365 spo sitedesign apply [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The ID of the site design to apply

`-u, --webUrl <webUrl>`
: The URL of the site to apply the site design to

`--asTask`
: Apply site design as task. Required for large site designs

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Apply the site design with ID 9b142c22-037f-4a7f-9017-e9d8c0e34b98 to the site collection https://contoso.sharepoint.com/sites/project-x

```sh
m365 spo sitedesign apply --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --webUrl https://contoso.sharepoint.com/sites/project-x
```

Apply large site design to the specified site

```sh
m365 spo sitedesign apply --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --webUrl https://contoso.sharepoint.com/sites/project-x --asTask
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)