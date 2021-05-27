# spo file list

Gets all files within the specified folder and site

## Usage

```sh
m365 spo file list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder from which to retrieve files is located

`-f, --folder <folder>`
: The server- or site-relative URL of the folder from which to retrieve files

`-r, --recursive`
: Switch to indicate whether the files should be returned recursively from folder structure.


--8<-- "docs/cmd/_global.md"

## Examples

Return all files from folder _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file list --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents'
```



Return all files from folder _Shared Documents_ and all the folders under _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_
```sh
m365 spo file list --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --recursive
```

Example : For the folder structure given below, using the ``--recursive`` option will return all the files : 
- AboutPnP.docx
- PnPTeamMembers.docx
- AboutO365Cli.docx
- O365-CLI-Contributors.docx
- command1-docs.docx
- command2-docs.docx
- AboutPnPJs.docx


```
_Shared Documents_
│   AboutPnP.docx
│   PnPTeamMembers.docx    
│
└───Office365CLI
│   │   AboutO365Cli.docx
│   │   O365-CLI-Contributors.docx
│   │
│   └───Documentation
│       │   command1-docs.docx
│       │   command2-docs.docx
│   
└───PnPJS
    │   AboutPnPJs.docx
```



