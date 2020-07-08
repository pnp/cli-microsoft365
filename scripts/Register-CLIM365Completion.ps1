function CLIMicrosoft365Completion {
  param($commandName, $wordToComplete, $cursorPosition)

  $commands = Get-Content $(Join-Path $PSScriptRoot ".." "commands.json" -Resolve) | ConvertFrom-Json
  $command = $commands
  $parent = $commands
  $replies = @{ }
  
  # split what's been typed by the user into words. First word is m365|microsoft365
  # which we can skip
  [string[]]$allWords = $wordToComplete.ToString().Split(" ", [StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Skip 1

  # if nothing was typed yet, return all top-level commands
  if ($null -eq $allWords -or $allWords.Count -eq 0) {
    $replies = $parent.psobject.properties | ForEach-Object { $_.Name }
    $replies | sort
    return
  }

  # number of the first word in the array that hasn't been matched with a command
  $wordNotMatched = 0
  do {
    $word = $allWords[$wordNotMatched]
    if ($word.StartsWith("-") -eq $true) {
      break
    }

    $parentCollection = if ($parent.Value) { $parent.Value.psobject.properties } else { $parent.psobject.properties }
    $parent = $parentCollection | Where-Object { $_.Name -eq $word }
    if ($null -ne $parent) {
      $command = $parent
      $wordNotMatched++
    }
    else {
      break
    }
  } until ($wordNotMatched -gt $allWords.Count - 1)

  $collection = if ($command.Value) { $command.Value.psobject.properties } else { $command.psobject.properties }
  $replies = $collection | ForEach-Object {
    $_.Name
  }

  # check if we matched the whole string or not, if we haven't we need to filter
  # the list of suggestions with only those that match the last partial string
  if ($wordNotMatched -lt $allWords.Length) {
    # filter the list of suggested commands only if the first unmatched word is
    # not an option
    $notMatchedWord = $allWords[$wordNotMatched]
    if ($notMatchedWord.StartsWith("-") -ne $true) {
      # since we didn't match the whole string, let's trim suggestions to match
      # the partial string, eg. `spo site l` > list
      $replies = $replies | Where-Object { $_ -Like "$($allWords[$wordNotMatched])*" }
    }
  }

  # check if the last word is an option
  $lastWord = $allWords[$allWords.Count - 1]
  if ($lastWord.StartsWith("-") -eq $true) {
    $option = $command.Value.psobject.properties | Where-Object { $_.Name -eq $lastWord }
    if ($null -ne $option) {
      # the option was matched. if the option is an enum, replace replies with
      # enum values
      if ($option.Value -is [System.Array]) {
        $replies = $option.Value | ForEach-Object { $_ }
      }
    }
    else {
      # check if there is a partial match
      $replies = $replies | Where-Object { $_ -Like "$($lastWord)*" }
    }
  }
  else {
    # check if the next to last word is an option
    $nextToLastWord = $allWords[$allWords.Count - 2]
    if ($nextToLastWord.StartsWith("-") -eq $true) {
      $option = $command.Value.psobject.properties | Where-Object { $_.Name -eq $nextToLastWord }
      if ($null -ne $option) {
        # the option was matched. if the option is an enum, check if the last word
        # fully matches one of the values. If it doesn't replace replies with
        # enum values
        if ($option.Value.Contains($lastWord) -ne $true) {
          $replies = $option.Value | Where-Object { $_ -Like "$($lastWord)*" }
        }
      }
    }
  }

  # remove used options
  $replies = $replies | Where-Object { $allWords.Contains($_) -ne $true }

  $replies | sort
}

Register-ArgumentCompleter -Native -CommandName @("m365", "microsoft365") -ScriptBlock $function:CLIMicrosoft365Completion