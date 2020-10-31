# Remove orphaned redirect sites

Author: [Albert-Jan Schot](https://www.cloudappie.nl/remove-orphaned-redirectsites/)

Changing the URL of a site results in a new site type: a Redirect Site. However this redirect site does not get removed if you delete the newly renamed site. This could result in orphaned redirect site collections that redirect to nothing. This script provides you with an overview of all orphaned redirect sites and allows you to quickly delete them.

```powershell tab="PowerShell Core"
$sites = m365 spo site classic list --t "REDIRECTSITE#0" --output json | ConvertFrom-Json

$sites | ForEach-Object {
  Write-Host -f Green "Processing redirect site: " $_.Url
  $siteUrl = $_.Url

  $redirectSite = Invoke-WebRequest -Uri $_.Url -MaximumRedirection 0 -SkipHttpErrorCheck
  $body = $null
  $siteUrl = $_.Url

  if($redirectSite.StatusCode -eq 308) {
    Try {
      [string]$newUrl = $redirectSite.Headers.Location;
      Write-Host -f Green " Redirects to: " $newUrl
      $body = Invoke-WebRequest -Uri $newUrl -SkipHttpErrorCheck
    }
    Catch{
     Write-Host $_.Exception
    }
    Finally {
      If($body.StatusCode -eq "200"){
       Write-host -f Yellow "  Target location still exists"
      }
      If($body.StatusCode -eq "404"){
        Write-Host -f Red "  Target location no longer exists, should be removed"
        m365 spo site remove --url $siteUrl
      }
    }
  }
}
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/
clear
sitestoremove=()
while read site; do
 siteUrl=$(echo ${site} | jq -r '.Url')
 echo "Checking old URL for redirect" $siteUrl
 redirectsite=$(curl -I -L --max-redirs 0 --silent --user-agent "Mozilla/5.0 (X11; Linux x86_64; rv:58.0) Gecko/20100101 Firefox/58.0" $siteUrl | sed -En 's/^location: (.*)/\1/p')

 echo "Redirects to" $redirectsite
 redirect_status=$(curl --write-out %{http_code} --silent --output /dev/null --user-agent "Mozilla/5.0 (X11; Linux x86_64; rv:58.0) Gecko/20100101 Firefox/58.0" ${redirectsite%$'\r'})

 if [[ "$redirect_status" -ne 404 ]] ; then
   echo "URL exists: $redirectsite"
 else
   echo "URL does not exist: $redirectsite"
   sitestoremove+=("$siteUrl")
 fi

done < <(m365 spo site classic list --t "REDIRECTSITE#0" -o json | jq -c '.[]')

if [ ${#sitestoremove[@]} = 0 ]; then
  exit 1
fi

printf '%s\n' "${sitestoremove[@]}"s
echo "Press Enter to start deleting (CTRL + C to exit)"
read foo

for site in "${sitestoremove[@]}"; do
   siteUrl=$(echo ${site} | jq -r '.Url')
  echo "Deleting site..."
  echo $siteUrl
   m365 spo site classic remove --url $siteUrl
done
```

Keywords:

- SharePoint Online
- Redirect sites
