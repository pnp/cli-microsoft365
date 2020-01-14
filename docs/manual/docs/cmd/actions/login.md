# GitHub Action - action-cli-login 
GitHub action to login to a tenant using Office 365 CLI.

This GitHub Action uses the [login command](https://pnp.github.io/office365-cli/cmd/login), to allow you log in to Office 365.

## Usage
### Pre-requisites
Create a workflow `.yml` file in your `.github/workflows` directory. An [example workflow](#example-workflow---office-365-cli-login) is available below. For more information, reference the GitHub Help Documentation for [Creating a workflow file](https://help.github.com/en/articles/configuring-a-workflow#creating-a-workflow-file).

### Inputs
- `ADMIN_USERNAME` : **Required** Username (email address of the admin)
- `ADMIN_PASSWORD` : **Required** Password of the admin

#### Optional requirement
Since this action requires user name and password which are sensitive pieces of information, it would be ideal to store them securely. We can achieve this in a GitHub repo by using [secrets](https://help.github.com/en/actions/automating-your-workflow-with-github-actions/creating-and-using-encrypted-secrets). So, click on `settings` tab in your repo and add 2 new secrets:
- `adminUsername` - store the admin user name in this (e.g. user@contoso.onmicrosoft.com)
- `adminPassword` - store the password of that user in this.
These secrets are encrypted and can only be used by GitHub actions. 

### Example workflow - Office 365 CLI Login
On every `push` build the code and then login to Office 365 before deploying.

```yaml
name: SPFx CICD with O365 CLI

on: [push]

jobs:
  build:
    ##
    ## Build code omitted
    ##
        
  deploy:
    needs: build
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node-version: [10.x]
    
    steps:
    
    ##
    ## Code to get the package omitted
    ##

    # Office 365 cli login action
    - name: Login to tenant
      uses: pnp/action-cli-login@v1
      with:
        ADMIN_USERNAME:  ${{ secrets.adminUsername }}
        ADMIN_PASSWORD:  ${{ secrets.adminPassword }}
    
    ##
    ## Code to deploy the package to tenant omitted
    ##
```