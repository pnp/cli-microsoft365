# Delete all (non-group connected) modern SharePoint sites

Author: [Laura Kokkarinen](https://laurakokkarinen.com/does-it-spark-joy-powershell-scripts-for-keeping-your-development-environment-tidy-and-spotless/#delete-all-non-group-connected-modern-sharepoint-sites)

When you delete Microsoft 365 groups, the modern group-connected team sites get deleted with them. The script below handles the remaining modern sites: communication sites and groupless team sites.

!!! attention
    There is a known issue running this script using PowerShell Core on macOS, see issue [#1266](https://github.com/pnp/cli-microsoft365/issues/1266) for further detail

```powershell tab="PowerShell Core"
$sparksjoy = "Cat Lovers United", "Extranet", "Hub"
$sites = m365 spo site classic list -o json |ConvertFrom-Json
$sites = $sites | where {  $_.template -eq "SITEPAGEPUBLISHING#0" -or $_.template -eq "STS#3" -and -not ($sparksjoy -contains $_.Title)}
if ($sites.Count -eq 0) { break }
$sites | Format-Table Title, Url, Template
Read-Host -Prompt "Press Enter to start deleting (CTRL + C to exit)"
$progress = 0
$total = $sites.Count
foreach ($site in $sites)
{
    $progress++
    write-host $progress / $total":" $site.Title
    write-host $site.Url
    m365 spo site classic remove --url $site.Url
}
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

sparksjoy=("Communication site" "Comm Site" "Hub")
sitestoremove=()
while read site; do
 siteTitle=$(echo ${site} | jq -r '.Title')
 echo $siteTitle
  exists=true
  for keep in "${sparksjoy[@]}"; do
    echo $keep
    if [ "$keep" == "$siteTitle" ] ; then
    echo "matched"
      exists=false
      break
    fi
  done
  if [ "$exists" = true ]; then
    sitestoremove+=("$site")
  fi

done < <(m365 spo site classic list -o json | jq -c '.[] | select(.Template == "SITEPAGEPUBLISHING#0" or .Template == "STS#3")')

if [ ${#sitestoremove[@]} = 0 ]; then
  exit 1
fi

printf '%s\n' "${sitestoremove[@]}"
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
- Microsoft 365 groups
