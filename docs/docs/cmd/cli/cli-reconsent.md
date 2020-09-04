# cli reconsent

Returns Azure AD URL to open in the browser to re-consent CLI for Microsoft 365 permissions

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

Get the URL to open in the browser to re-consent CLI for Microsoft 365 permissions

```sh
cli reconsent
```

## More information

- Re-consent the PnP Microsoft 365 Management Shell Azure AD application: [https://pnp.github.io/cli-microsoft365/user-guide/connecting-office-365/#re-consent-the-pnp-office-365-management-shell-azure-ad-application](https://pnp.github.io/cli-microsoft365/user-guide/connecting-office-365/#re-consent-the-pnp-office-365-management-shell-azure-ad-application)
