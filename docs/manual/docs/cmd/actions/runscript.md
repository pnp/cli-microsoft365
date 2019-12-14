# GitHub Action - action-cli-runscript
GitHub action to run a script using the Office 365 CLI

This GitHub Action uses [Office 365 CLI](https://pnp.github.io/office365-cli/), to run a line of script supplied to it or run code in a script file supplied to it.

## Usage
### Pre-requisites
Create a workflow `.yml` file in `.github/workflows` directory of your repo. An [example workflow](#example-workflow---office-365-cli-runscript) is available below. For more information, reference the GitHub Help Documentation for [Creating a workflow file](https://help.github.com/en/articles/configuring-a-workflow#creating-a-workflow-file).

### Inputs
- `O365_CLI_SCRIPT` : The script to run
- `O365_CLI_SCRIPT_PATH` : Relative path of the script in your repo.

One of `O365_CLI_SCRIPT_PATH` / `O365_CLI_SCRIPT` is mandatory, in case both are defined `O365_CLI_SCRIPT_PATH` gets preference.

### Example workflow - Office 365 CLI Runscript
On every `push` build the code, then deploy and then send an email using Office 365 CLI Runscript action.

```yaml
name: SPFx CICD with O365 CLI

on: [push]

jobs:
  build:
    ##
    ## Build code omitted
    ##
        
  deploy:
    ##
    ## Code to deploy the package to tenant omitted
    ##

  runscript:
    
    # Office 365 CLI runscript action option 1 (single line of script as input)
    - name: Send email
      uses: pnp/action-cli-runscript@v1
      env:
        O365_CLI_SCRIPT: o365 spo mail send --webUrl https://contoso.sharepoint.com/sites/teamsite --to 'user@contoso.onmicrosoft.com' --subject 'Deployment done' --body '<h2>Office 365 CLI</h2> <p>The deployment is complete.</p> <br/> Email sent via Office 365 CLI GitHub Action.'
    
    # Office 365 CLI runscript action option 2 (script file as input)
    - name: Create lists
      uses: pnp/action-cli-runscript@v1
      env:
        O365_CLI_SCRIPT_PATH: /script/lists.sh 
        #lists.sh will have all the required Office 365 CLI commands
```