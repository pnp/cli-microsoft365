# Analyze users for known data breaches with have i been pwned

Inspired by: [Albert-Jan Schot](https://www.cloudappie.nl/cli-microsoft-haveibeenpwned-status/)

Validate all your users against known breaches with the have i been pwned api. That way you can quickly scan if your users are part of any known breaches.

=== "PowerShell"

    ```powershell
    $apiKey = "<PUTYOURKEYHERE>"
    $m365Status = m365 status

    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    $users = m365 aad user list --properties "displayName,userPrincipalName" | ConvertFrom-Json

    $users | ForEach-Object {
      $user = $_
      $i++
      Write-Host "Check HBIP status for user '$($user.userPrincipalName)' - ($i/$($users.length))"

      $hbipStatus = m365 aad user hibp --userName $user.userPrincipalName --apiKey $apiKey --verbose | ConvertFrom-Json

      if ($hbipStatus -ne "No pwnage found") {
        Write-Host -ForegroundColor Red "Issue with user '$($user.userPrincipalName)'"
        $hbipStatus
      }

      Start-Sleep -Milliseconds 1500
    }
    ```

Keywords:

- Azure
- Microsoft 365
- PowerShell
- Security
