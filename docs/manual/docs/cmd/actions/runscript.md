# runscript action

This action runs an Office 365 CLI command or a set of commands.

## Inputs

### `O365_CLI_SCRIPT`
The script to run

### `O365_CLI_SCRIPT_PATH`
Relative path of the script in your repo.

One of O365_CLI_SCRIPT_PATH / O365_CLI_SCRIPT is mandatory, in case both are defined O365_CLI_SCRIPT_PATH gets preference.

## Usage

```sh
uses: pnp/office365-cli/actions/runscript@master
      env:
        O365_CLI_SCRIPT: o365 spo mail send --webUrl https://contoso.sharepoint.com/sites/teamsite --to 'user@contoso.onmicrosoft.com' --subject 'Deployment done' --body '<h2>Office 365 CLI</h2> <p>The deployment is complete.</p> <br/> Email sent via Office 365 CLI GitHub Action.'
```

```sh
uses: pnp/office365-cli/actions/runscript@master
      env:
        O365_CLI_SCRIPT_PATH: /script/lists.sh
```