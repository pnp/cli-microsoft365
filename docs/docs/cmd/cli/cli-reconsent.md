# cli reconsent

Returns Azure AD URL to open in the browser to re-consent CLI for Microsoft 365 permissions

## Usage

```sh
m365 cli reconsent [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get the URL to open in the browser to re-consent CLI for Microsoft 365 permissions

```sh
m365 cli reconsent
```

## Response

=== "JSON"

    ```json
    "To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent"
    ```

=== "Text"

    ```text
    To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent
    ```

=== "CSV"

    ```csv
    To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to https://login.microsoftonline.com/common/oauth2/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&prompt=admin_consent
    ```

## More information

- Re-consent the PnP Microsoft 365 Management Shell Azure AD application: [https://pnp.github.io/cli-microsoft365/user-guide/connecting-office-365/#re-consent-the-pnp-office-365-management-shell-azure-ad-application](https://pnp.github.io/cli-microsoft365/user-guide/connecting-office-365/#re-consent-the-pnp-office-365-management-shell-azure-ad-application)
