# GitHub Actions workflow

This file shows below and example of a workflow that can be used in a SPFx GitHub repo. Clicking on 'Actions' in a GitHub repo opens a YAML file. In that file the following code can be entered. 

The code does the following
* Builds the project
* Bundles the solution
* Creates the package
* Uploads it to the artifacts
* Downloads the package from artifacts
* Uses Office CLI Login action to login to the tenant
* Deploys the package using Office 365 CLI deploy action
* Sends an email using the Office 365 CLI runscript action

```
name: SPFx CICD with O365 CLI

on: [push]

jobs:
  build:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node-version: [10.x]
    
    steps:
    # Checkout code
    - name: Checkout code
      uses: actions/checkout@v1
      
    # Setup node.js runtime
    - name: Use Node.js ${{ matrix.node-version }}
      uses: actions/setup-node@v1
      with:
        node-version: ${{ matrix.node-version }}
    
    # npm install
    - name: Run npm install
      run: npm install
    
    # gulp bundle and package solution
    - name: Bundle and package
      run: |
        gulp bundle --ship
        gulp package-solution --ship    
    
    # Upload artifacts (sppkg file)
    - name: Upload artifact (sppkg file)
      uses: actions/upload-artifact@v1.0.0
      with:
        name: output
        path: sharepoint/solution/${{ secrets.sppkgFileName }}
        
  deploy:
    needs: build
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node-version: [10.x]
    
    steps:
    
    # Download package (sppkg file)
    - name: Download pacakge (sppkg file)
      uses: actions/download-artifact@v1
      with:
        name: output
    
    - name: Office 365 CLI Login
      uses: pnp/office365-cli/actions/login@master
      env:
        ADMIN_USERNAME:  ${{ secrets.adminUsername }}
        ADMIN_PASSWORD:  ${{ secrets.adminPassword }}
    - name: Office 365 CLI Deploy
      uses: pnp/office365-cli/actions/deploy@master
      env:
        APP_FILE_PATH: output/${{ secrets.sppkgFileName }}
    - name: Office 365 CLI Send email
      uses: pnp/office365-cli/actions/runscript@master
      env:
         O365_CLI_SCRIPT: o365 spo mail send --webUrl https://contoso.sharepoint.com/sites/teamsite --to 'user@contoso.onmicrosoft.com' --subject 'Deployment done' --body '<h2>Office 365 CLI</h2> <p>The deployment is complete.</p> <br/> Email sent via Office 365 CLI GitHub Action.'
```