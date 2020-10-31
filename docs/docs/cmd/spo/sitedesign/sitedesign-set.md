# spo sitedesign set

Updates a site design with new values

## Usage

```sh
m365 spo sitedesign set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The ID of the site design to update

`-t, --title [title]`
: The new display name of the updated site design

`-w, --webTemplate [webTemplate]`
: The new template to add the site design to. Allowed values `TeamSite,CommunicationSite`

`-s, --siteScripts [siteScripts]`
: Comma-separated list of new site script IDs. The scripts will run in the order listed

`-d, --description [description]`
: The new display description of updated site design

`-m, --previewImageUrl [previewImageUrl]`
: The new URL of a preview image. If none is specified SharePoint will use a generic image

`-a, --previewImageAltText [previewImageAltText]`
: The new alt text description of the image for accessibility

`-v, --version [version]`
: The new version number for the site design

`--isDefault [isDefault]`
: Set to true if the site design is applied as the default site design

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If you had previously set the `isDefault` option to `true`, and wish for it to remain `true`, you must pass in this option again or it will be reset to `false`.

When specifying IDs of site scripts to use with your site design, ensure that the IDs refer to existing site scripts or provisioning sites using the design will lead to unexpected results.

## Examples

Update the site design title and version

```sh
m365 spo sitedesign set --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --title "Contoso site design" --version 2
```

Update the site design to be the default design for provisioning modern communication sites

```sh
m365 spo sitedesign set --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --webTemplate CommunicationSite  --isDefault true
```


Update the site design to be the default design for provisioning modern communication sites, with specific scripts

```sh
m365 spo sitedesign set --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --webTemplate CommunicationSite  --isDefault true --siteScripts "19b0e1b2-e3d1-473f-9394-f08c198ef43e,b2307a39-e878-458b-bc90-03bc578531d6"
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
- Customize a default site design: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/customize-default-site-design](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/customize-default-site-design)
- Site design JSON schema: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema)
