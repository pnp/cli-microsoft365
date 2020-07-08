# Authenticate with and call the Microsoft Graph 

Author: [Garry Trinder](https://github.com/garrytrinder)

Obtain a new access token for the Microsoft Graph and use it an HTTP request.

```powershell tab="PowerShell Core"
$token = m365 util accesstoken get --resource https://graph.microsoft.com --new
$me = Invoke-RestMethod -Uri https://graph.microsoft.com/v1.0/me -Headers @{"Authorization"="Bearer $token"}
$me
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

token=`m365 util accesstoken get --resource https://graph.microsoft.com --new`
me=`curl https://graph.microsoft.com/v1.0/me -H "Authorization: Bearer $token"`
echo $me | jq
```

Keywords:

- Microsoft Graph
- Access Token
- HTTP
