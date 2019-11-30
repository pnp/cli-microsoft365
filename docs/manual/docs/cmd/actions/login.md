# login action

This action helps to login into a tenant.

## Inputs

### `ADMIN_USERNAME`
**Required** Username (email address of the admin)

### `ADMIN_PASSWORD`
**Required** Password of the admin

## Usage

```sh
uses: pnp/office365-cli/actions/login@master
      env:
        ADMIN_USERNAME:  ${{ secrets.adminUsername }}
        ADMIN_PASSWORD:  ${{ secrets.adminPassword }}
```