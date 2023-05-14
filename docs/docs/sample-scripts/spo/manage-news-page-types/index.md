---
tags:
  - pages
  - libraries
---

# Create custom views to differentiate SharePoint news page types

Author: [Nanddeep Nachan](https://github.com/nanddeepn), Inspired by [João Ferreira](https://sharepoint.handsontek.net/2020/08/23/effectively-manage-sharepoint-news-part-1/)

SharePoint stores the news in the Site Pages library along with all the other pages, you can easily end up with hundreds of pages in the library with no easy way to identify Pages, Spaces, News and News Links.

The following script shows how to create custom views to differentiate News types in Site Pages library.

=== "PowerShell"

    ```powershell
    param
    (
        [Parameter(Mandatory = $true, HelpMessage="URL of the site where the list is located")][string] $WebUrl,
        [Parameter(Mandatory = $false, HelpMessage="Title of the list to which the view should be added")][string] $ListTitle = "Site Pages"
    )

    try {
      $m365Status = m365 status
      if ($m365Status -eq "Logged Out") {
        Write-Host "Logging in the User!"
        m365 login --authType browser
      }

      Write-Host "Creating view - All News"
      m365 spo list view add --webUrl $WebUrl --listTitle $ListTitle --title "All News" --fields "Title,Name,Editor,Modified" --viewQuery "<Where><Eq><FieldRef Name='PromotedState'></FieldRef><Value Type='Number'>2</Value></Eq></Where>" --paged

      Write-Host "Creating view - SharePoint News"
      m365 spo list view add --webUrl $WebUrl --listTitle $ListTitle --title "SharePoint News" --fields "Title,Name,Editor,Modified" --viewQuery "<Where><And><Eq><FieldRef Name='PromotedState' /><Value Type='Number'>2</Value></Eq><Eq><FieldRef Name='PageLayoutType' /><Value Type='Text'>Article</Value></Eq></And></Where>" --paged

      Write-Host "Creating view - News Link"
      m365 spo list view add --webUrl $WebUrl --listTitle $ListTitle --title "News Link" --fields "Title,Name,Editor,Modified" --viewQuery "<Where><And><Eq><FieldRef Name='PromotedState' /><Value Type='Number'>2</Value></Eq><Eq><FieldRef Name='PageLayoutType' /><Value Type='Text'>RepostPage</Value></Eq></And></Where>" --paged
    }
    catch {
        Write-Host -f Red "Error generating test documents: " $_.Exception.Message
    }

    Write-Host "Finished"
    ```
