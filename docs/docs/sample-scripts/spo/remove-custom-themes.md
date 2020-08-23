# Delete custom color themes from SharePoint

Author: [Laura Kokkarinen](https://laurakokkarinen.com/does-it-spark-joy-powershell-scripts-for-keeping-your-development-environment-tidy-and-spotless/#delete-all-custom-color-themes-from-sharepoint)

Have you been creating a lot of beautiful themes lately and testing them in your dev tenant, but don’t want to keep them anymore? If yes, then this PowerShell script is for you.

```powershell tab="PowerShell Core"
$sparksjoy = "Cat Lovers United", "Multicolored theme"
$themes = m365 spo theme list -o json | ConvertFrom-Json
$themes = $themes | where {-not ($sparksjoy -contains $_.name)}
$themes | Format-Table name
if ($themes.Count -eq 0) { break }
Read-Host -Prompt "Press Enter to start deleting (CTRL + C to exit)"
$progress = 0
$total = $themes.Count
foreach ($theme in $themes)
{
  $progress++
  write-host $progress / $total":" $theme.name
  m365 spo theme remove --name "$($theme.name)" --confirm
}
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

sparksjoy=("Cat Lovers United" "Multicolored theme")
themestoremove=()
while read theme; do
  exists=false
  for keep in "${sparksjoy[@]}"; do
    if [ "$keep" == "$theme" ] ; then
      exists=true
      break
    fi
  done
  if [ "$exists" = false ]; then
    themestoremove+=("$theme")
  fi
done < <(m365 spo theme list -o json | jq -r '.[].name')

if [ ${#themestoremove[@]} = 0 ]; then
  exit 1
fi

printf '%s\n' "${themestoremove[@]}"
echo "Press Enter to start deleting (CTRL + C to exit)"
read foo

for theme in "${themestoremove[@]}"; do
  echo "Deleting $theme..."
  m365 spo theme remove --name "$theme" --confirm
done
```

Keywords:

- SharePoint Online
- Themes
