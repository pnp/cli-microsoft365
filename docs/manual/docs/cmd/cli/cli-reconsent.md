# cli reconsent

Returns Azure AD URL to open in the browser to re-consent Office 365 CLI permissions

## Usage

```sh
cli reconsent [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get the URL to open in the browser to re-consent Office 365 CLI permissions

```sh
cli reconsent
```

## More information

- Re-consent the PnP Office 365 Management Shell Azure AD application: [https://pnp.github.io/office365-cli/user-guide/connecting-office-365/#re-consent-the-pnp-office-365-management-shell-azure-ad-application](https://pnp.github.io/office365-cli/user-guide/connecting-office-365/#re-consent-the-pnp-office-365-management-shell-azure-ad-application)
