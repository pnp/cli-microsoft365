$location = Get-Location
$files = Get-ChildItem $location -Recurse -Include *.md
$counter = 0
$pathNames = @()

function ReplaceAdmonition ([string[]]$content, [string]$marker, [string]$replaceMarker) {
  if ($content.Contains($marker)) {
    for ($i = 0; $i -lt $content.Length; $i++) {
      if ($content[$i].ToLower().Contains($marker)) {
        $content[$i] = $content[$i] -replace $marker, $replaceMarker
        $content[$i+1] = $content[$i+1] -replace '    ', ''
        $content[$i+1] += "`n:::"
      }
    }
  }

  return $content
}

function UpdateImports ([string[]]$content, [string]$import) {
  $frontMatterEndSyntax = $content | Select-String -Pattern '^---$' | Select-Object -Last 1

  if ($frontMatterEndSyntax) {
    $frontMatterLastLineNumber = $frontMatterEndSyntax.LineNumber
    $content[$frontMatterLastLineNumber] += "`n$($import)"
  } else {
    $content = ,"$($import)" + $content
  }

  $content | Out-File "temp.txt"
  $updatedContent = Get-Content "temp.txt"
  Remove-Item -LiteralPath "temp.txt"

  $lastImportLineNumber = ($updatedContent | Select-String -Pattern "^import.*;`$" | Select-Object -Last 1).LineNumber

  if ($lastImportLineNumber -and !([string]::IsNullOrWhiteSpace($updatedContent[$lastImportLineNumber]))) {
    $updatedContent[$lastImportLineNumber] = "`n$($updatedContent[$lastImportLineNumber])"
  }

  return $updatedContent
}

foreach ($file in $files) {
  $counter++
  Write-Progress -Activity 'Processing files' -Status $file.VersionInfo.FileName -PercentComplete (($counter / $files.count) * 100)

  $content = Get-Content -path $file.VersionInfo.FileName

  #* Update code tabs
  $frontMatterLastLineNumber

  if($content | Select-String '=== "', -Quiet) {
    $content = UpdateImports $content "import Tabs from '@theme/Tabs';`nimport TabItem from '@theme/TabItem';"

    $started = $false
    $language = ""

    for ($i = 0; $i -lt $content.Length; $i++) { 
      if ($content[$i] -match '=== "') {
        $language = ($content[$i] -split '"')[1]
        $content = $content | Skip-Object -Index $i
        $openingTag = ""

        if (!$started) {
          $openingTag = "<Tabs>`n"
          $started = $true
        }

        $content[$i] = "$($openingTag)  <TabItem value=`"$($language)`">`n"
      }
      
      elseif ($started -and $content[$i].EndsWith('```')) {
        $content[$i] = $content[$i].Substring(2) + "`n`n  </TabItem>"
        $content = $content | Skip-Object -Index ($i+1)

        if (!($content[$i+1] -match '=== "')) {
          $started = $false;
          $content[$i] += "`n</Tabs>`n";
        }
      }

      elseif ($started -and $content[$i].Length -ge 2) {
        $content[$i] = $content[$i].Substring(2);
      }
    }
  }

  #* Update definition-lists

  if($content -contains '## Options') {
    for ($i = 0; $i -lt $content.Length; $i++) {
      if ($content[$i] -match '## Options') {
        $content[$i+1] = "`n``````md defintion-list"
      }

      if ($content[$i] -match '--8<-- "docs/cmd/_global.md"') {
        $content[$i-1] = "```````n"
      }
    }
  }

  #* Update globals & CLISettings

  if($content -contains '--8<-- "docs/cmd/_global.md"') {    
    $content = UpdateImports $content "import Global from '/docs/cmd/_global.mdx';"
    $content = $content -replace '--8<-- "docs/cmd/_global.md"', '<Global />'
  }

  if($content -contains '--8<-- "docs/_clisettings.md"') {    
    $content = UpdateImports $content "import CLISettings from '/docs/_clisettings.mdx';"
    $content = $content -replace '--8<-- "docs/_clisettings.md"', '<CLISettings />'
  }

  #* Update admonitions

  $content = ReplaceAdmonition $content '!!! attention' ':::caution'
  $content = ReplaceAdmonition $content '!!! note' ':::note'
  $content = ReplaceAdmonition $content '!!! important' ':::info'
  $content = ReplaceAdmonition $content '!!! tip' ':::tip'
  $content = ReplaceAdmonition $content '!!! warning' ':::danger'

  Set-Content -Path $file.VersionInfo.FileName -Value $content

  $newName = $file.Name -replace ".md", ".mdx"
  $pathNames += @{OG = $file.Name; New = $newName}
  Rename-Item -Path $file.FullName -NewName $newName
}

$pathNames = $pathNames  | Group-Object OG, New | ForEach-Object { $_.Group | Select-Object -First 1 }

Write-Progress -Activity 'Processing files' -Status "Ready" -Completed

$files = Get-ChildItem $location -Recurse -Include ('*.md', '*.mdx')
$counter = 0
foreach ($file in $files) {
  $counter++
  Write-Progress -Activity 'Processing files' -Status $file.VersionInfo.FileName -PercentComplete (($counter / $files.count) * 100)
  
  $content = Get-Content -Path $file.VersionInfo.FileName

  foreach ($pathName in $pathNames) {
    $content = $content -replace $pathName.OG, $pathName.New
  }
  
  Set-Content -Path $file.VersionInfo.FileName -Value $Content
}

Write-Progress -Activity 'Update links with mdx' -Status "Ready" -Completed