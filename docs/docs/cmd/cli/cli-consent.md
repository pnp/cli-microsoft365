# cli consent

Consent additional permissions for the Azure AD application used by the CLI for Microsoft 365

## Usage

```sh
m365 cli consent [options]
```

## Options

`-s, --service <service>`
: Service for which to consent permissions. Allowed values: `yammer`

--8<-- "docs/cmd/_global.md"

## Remarks

Using the `cli consent` command you can consent additional permissions for the Azure AD application used by the CLI for Microsoft 365. This is for example necessary to use Yammer commands, which require the Yammer API permission that isn't granted to the CLI by default.

After executing the command, the CLI for Microsoft 365 will present you with a URL that you need to open in the web browser in order to consent the permissions for the selected Microsoft 365 service.

To simplify things, rather than wondering which permissions you should grant for which CLI commands, this command allows you to easily grant all the necessary permissions for using commands for the specified Microsoft 365 service, like Yammer.

## Examples

Consent permissions to the Yammer API

```sh
m365 cli consent --service yammer
```
