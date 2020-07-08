. $(Join-Path . Register-CLIM365Completion.ps1)

$tests = @{
  "m365" = @("aad","accesstoken","consent","flow","graph","help","login","logout","onedrive","outlook","pa","planner","spfx","spo","status","teams","tenant","yammer");
  "m365 " = @("aad","accesstoken","consent","flow","graph","help","login","logout","onedrive","outlook","pa","planner","spfx","spo","status","teams","tenant","yammer");
  "m365 s" = @("spfx","spo","status");
  "m365 spo" = @("app","apppage","cdn","contenttype","contenttypehub","customaction","externaluser","feature","field","file","folder","get","hidedefaultthemes","homesite","hubsite","list","listitem","mail","navigation","orgassetslibrary","orgnewssite","page","propertybag","report","search","serviceprincipal","set","site","sitedesign","sitescript","sp","storageentity","tenant","term","theme","web");
  "m365 spo site" = @("add","appcatalog","classic","commsite","get","groupify","inplacerecordsmanagement","list","rename","set");
  "m365 spo site list" = @("--debug","--filter","--help","--output","--type","--verbose","-f","-o");
  "m365 b" = $null
  "m365 spo site list -" = @("--debug","--filter","--help","--output","--type","--verbose","-f","-o");
  "m365 spo site list -b" = $null;
  "m365 spo site list -o" = @("json","text");
  "m365 spo site list -o j" = @("json");
  "m365 spo site list --o" = @("--output");
  "m365 spo site list -o json" = @("--debug","--filter","--help","--output","--type","--verbose","-f");
  "m365 spo site list --debug" = @("--filter","--help","--output","--type","--verbose","-f","-o");
}

$tests.Keys | ForEach-Object {
  Write-Host "$($_)..." -NoNewLine
  $completion = CLIMicrosoft365Completion "" $_ 6
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