# deploy action

This action adds, deploys and installs (only to site collection) an app.

## Inputs

### `APP_FILE_PATH`
**Required** Relative path of the app in your repo.

### `SCOPE`
Scope of the app catalog: `tenant|sitecollection`. Default `tenant`

### `SITE_COLLECTION_URL`
The URL of the site collection where the solution package will be added and installed. It must be specified when the scope is `sitecollection`

## Usage

```sh
uses: pnp/office365-cli/actions/deploy@master
      env:
        APP_FILE_PATH: sharepoint/solution/spfx-o365-cli-action.sppkg
```

```sh
uses: pnp/office365-cli/actions/deploy@master
      env:
        APP_FILE_PATH: sharepoint/solution/spfx-o365-cli-action.sppkg
        SCOPE: sitecollection
        SITE_COLLECTION_URL: https://contoso.sharepoint.com/sites/teamsite
```