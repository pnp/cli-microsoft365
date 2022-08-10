# List all checked out files in SharePoint

Author: [Veronique Lengelle](https://veronicageek.com/2020/find-checked-out-files-across-multiple-site-collections/)

## Checked out files in a specific site

This script will retrieve all the checked out files in a particular site.

=== "PowerShell"

    ```powershell
    $m365Status = m365 status --output text
    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    #Declare variables
    $siteURL = "<YOUR-SITE-URL>"
    $allLibs = m365 spo list list --webUrl $siteURL --query "[?BaseTemplate == ``101``]" -o json | ConvertFrom-Json
    $resultsForSite = @()

    foreach($library in $allLibs){
        $allDocs = m365 spo file list --webUrl $siteURL --folder $library.Url --recursive -o json | ConvertFrom-Json

        foreach($document in $allDocs){
            if($document.CheckOutType -eq [int64]0){
                $resultsForSite += [pscustomobject][ordered]@{
                    LibraryName = $library.Title
                    DocumentName = $document.Name
                    RelativePath = $document.ServerRelativeUrl
                }
            }
        }
    }
    $resultsForSite
    ```

## Checked out files for a specific document library on a site

This script will retrieve all the checked out files in a specific document library on a particular site.

=== "PowerShell"

    ```powershell
    $m365Status = m365 status --output text
    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    #Declare variables
    $siteURL = "<YOUR-SITE-URL>"
    $libraryName = "<LIBRARY NAME>"  ## Example: "Shared Documents"
    $allDocuments = m365 spo file list --webUrl $siteURL --folder $("$libraryName") --recursive -o json | ConvertFrom-Json
    $resultsForLib = @()

    #Loop through each document
    foreach($doc in $allDocuments){
        if($doc.CheckOutType -eq [int64]0){
            $resultsForLib += [pscustomobject][ordered]@{
                LibraryName = $libraryName
                DocName = $doc.Name
                RelativePath = $doc.ServerRelativeUrl
            }
        }
    }
    $resultsForLib
    ```

## Checked out files on each document library for multiple sites provided in a CSV file

This script will loop through each site from your CSV file, and retrieve all the checked out files from each document library. Your CSV file should contain a single header called "siteURL" with each URL per row:

| siteURL                                    |
| ------------------------------------------ |
| https://contoso.sharepoint.com/sites/site1 |
| https://contoso.sharepoint.com/sites/site2 |
| https://contoso.sharepoint.com/sites/site3 |

!!! important
    Depending on the number of sites in your .csv file, the number of libraries as well as the number of files, the below script can take a very long time to provide results.

=== "PowerShell"

    ```powershell
    $m365Status = m365 status --output text
    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    #Declare variables
    $allSites = Import-Csv -Path "<YOUR-FILE-PATH>"
    $resultsForEachSC = @()

    foreach($row in $allsites){
        #Get the libraries
        $allLibraries = m365 spo list list --webUrl $row.siteURL --query "[?BaseTemplate == ``101``]" -o json | ConvertFrom-Json

        foreach($lib in $allLibraries){

            #Get all the documents
            $allDocs = m365 spo file list --webUrl $row.siteURL --folder $lib.Url --recursive -o json | ConvertFrom-Json

            foreach($docu in $allDocs){
                if($docu.CheckOutType -eq [int64]0){
                    $resultsForEachSC += [pscustomobject][ordered]@{
                        LibraryName = $lib.Title
                        DocumentName = $docu.Name
                        RelativePath = $docu.ServerRelativeUrl
                    }
                }
            }
        }
    }
    $resultsForEachSC
    ```

Keywords

- SharePoint Online
- Governance
