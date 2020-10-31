# spo sitedesign add

Adds site design for creating modern sites

## Usage

```sh
m365 spo sitedesign add [options]
```

## Options

`-h, --help`
: output usage information

`-t, --title <title>`
: The display name of the site design

`-w, --webTemplate <webTemplate>`
: Identifies which base template to add the design to. Allowed values `TeamSite,CommunicationSite`

`-s, --siteScripts <siteScripts>`
: Comma-separated list of site script IDs. The scripts will run in the order listed

`-d, --description [description]`
: The display description of site design

`-m, --previewImageUrl [previewImageUrl]`
: The URL of a preview image. If none is specified SharePoint will use a generic image

`-a, --previewImageAltText [previewImageAltText]`
: The alt text description of the image for accessibility

`--isDefault`
: Set if the site design is applied as the default site design

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Each time you execute the `spo sitedesign add` command, it will create a new site design with a unique ID. Before creating a site design, be sure that another design with the same name doesn't already exist.

When specifying IDs of site scripts to use with your site design, ensure that the IDs refer to existing site scripts or provisioning sites using the design will lead to unexpected results.

## Examples

Create new site design for provisioning modern team sites

```sh
m365 spo sitedesign add --title "Contoso team site" --webTemplate TeamSite --siteScripts "19b0e1b2-e3d1-473f-9394-f08c198ef43e,b2307a39-e878-458b-bc90-03bc578531d6"
```

Create new default site design for provisioning modern communication sites

```sh
m365 spo sitedesign add --title "Contoso communication site" --webTemplate CommunicationSite --siteScripts "19b0e1b2-e3d1-473f-9394-f08c198ef43e" --isDefault
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
- Customize a default site design: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/customize-default-site-design](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/customize-default-site-design)
- Site design JSON schema: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema)
