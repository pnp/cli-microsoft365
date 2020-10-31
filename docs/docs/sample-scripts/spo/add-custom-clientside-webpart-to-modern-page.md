# Add custom client-side web part to modern page

Author: [Yannick Plenevaux](https://ypcode.wordpress.com)

You've built an amazing new web part and now you want to programmatically add it to a modern page. This sample helps you add your web part to the page with your custom properties that might be dynamic according to your script.

```powershell tab="PowerShell Core"
$site = "https://contoso.sharepoint.com/sites/site1"
$pageName = "AModernPage.aspx"
$webPartId = "af660fc1-c09b-4c15-b093-2b74b047286b"

$choice1 = "Choice 1"
$choice2 = "Choice 2"

# Put all the web part properties in a PowerShell hashtable
$webPartProps = @{
    myChoices              = @($choice1, $choice2);
    description            = 'My "Awesome" web part';
};

# Build JSON string from PowerShell hashtable object
$webPartPropsJson = $webPartProps | ConvertTo-Json -Compress
# Make sure to add the backticks, double the JSON double-quotes and escape double quotes in properties'values
$webPartPropsJson = '`"{0}"`' -f $webPartPropsJson.Replace('\','\\').Replace('"', '""')

m365 spo page clientsidewebpart add -u $site -n $pageName --webPartId $webPartId --webPartProperties $webPartPropsJson
```

```bash tab="Bash"
#!/bin/bash
site=https://contoso.sharepoint.com/sites/site1
pageName=AModernPage.aspx
webPartId=af660fc1-c09b-4c15-b093-2b74b047286b

choice1='Choice X'
choice2='Choice Z'
description='My "Super Awesome" web part';
# Build the JSON including your dynamic values with printf
# For each argument that might be dynamic, we escape the double quotes " with \"
# Make sure not to ommit the surrounding back ticks and surrounding double quotes for each arguments
printf -v webPartPropsJson '`{"myChoices":["%s","%s"], "description":"%s"}`' "${choice1//\"/\\\"}" "${choice2//\"/\\\"}" "${description//\"/\\\"}"

m365 spo page clientsidewebpart add -u $site -n $pageName --webPartId $webPartId --webPartProperties $webPartPropsJson
```

Keywords:

- SharePoint Online
- Client-side WebPart
- Modern page
- WebPart Properties
- JSON
