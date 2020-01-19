. $(Join-Path . Register-O365CLICompletion.ps1)

$tests = @{
  "o365" = @("aad","accesstoken","consent","flow","graph","help","login","logout","onedrive","outlook","pa","planner","spfx","spo","status","teams","tenant","yammer");
  "o365 " = @("aad","accesstoken","consent","flow","graph","help","login","logout","onedrive","outlook","pa","planner","spfx","spo","status","teams","tenant","yammer");
  "o365 s" = @("spfx","spo","status");
  "o365 spo" = @("app","apppage","cdn","contenttype","contenttypehub","customaction","externaluser","feature","field","file","folder","get","hidedefaultthemes","homesite","hubsite","list","listitem","mail","navigation","orgassetslibrary","orgnewssite","page","propertybag","report","search","serviceprincipal","set","site","sitedesign","sitescript","sp","storageentity","tenant","term","theme","web");
  "o365 spo site" = @("add","appcatalog","classic","commsite","get","groupify","inplacerecordsmanagement","list","rename","set");
  "o365 spo site list" = @("--debug","--filter","--help","--output","--type","--verbose","-f","-o");
  "o365 b" = $null
  "o365 spo site list -" = @("--debug","--filter","--help","--output","--type","--verbose","-f","-o");
  "o365 spo site list -b" = $null;
  "o365 spo site list -o" = @("json","text");
  "o365 spo site list -o j" = @("json");
  "o365 spo site list --o" = @("--output");
  "o365 spo site list -o json" = @("--debug","--filter","--help","--output","--type","--verbose","-f");
  "o365 spo site list --debug" = @("--filter","--help","--output","--type","--verbose","-f","-o");
}

$tests.Keys | ForEach-Object {
  Write-Host "$($_)..." -NoNewLine
  $completion = Office365Completion "" $_ 6
  if ($null -eq $completion -and $null -eq $tests.Item($_)) {
    Write-Host "PASSED" -ForegroundColor Green
  }
  elseif ([String]::Join(",", $completion) -eq [String]::Join(",", $tests.Item($_))) {
    Write-Host "PASSED" -ForegroundColor Green
  }
  else {
    Write-Host "FAILED" -ForegroundColor Red
    Write-Host "  Expected: $([String]::Join(",",$tests.Item($_)))"
    Write-Host "  Actual:   $([String]::Join(",", $completion))"
  }
}