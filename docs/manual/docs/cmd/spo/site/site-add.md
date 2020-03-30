# spo site add

Creates new modern site

## Usage

```sh
spo site add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--type [type]`|Type of modern sites to list. Allowed values `TeamSite,CommunicationSite`, default `TeamSite`
`-u, --url <url>`|Site URL (applies only to communication sites)
`-a, --alias <alias>`|Site alias, used in the URL and in the team site group e-mail (applies only to team sites)
`-t, --title <title>`|Site title
`-d, --description [description]`|Site description
`-c, --classification [classification]`|Site classification
`-l, --lcid [lcid]`|Site language in the LCID format, eg. _1033_ for _en-US_
`--isPublic`|Determines if the associated group is public or not (applies only to team sites)
`--shareByEmailEnabled`|Determines whether it's allowed to share file with guests (applies only to communication sites)
`--allowFileSharingForGuestUsers`|(deprecated. Use `shareByEmailEnabled` instead) Determines whether it's allowed to share file with guests (applies only to communication sites)
`--siteDesign [siteDesign]`|Type of communication site to create. Allowed values `Topic,Showcase,Blank`, default `Topic`. When creating a communication site, specify either `siteDesign` or `siteDesignId` (applies only to communication sites)
`--siteDesignId [siteDesignId]`|Id of the custom site design to use to create the site. When creating a communication site, specify either `siteDesign` or `siteDesignId` (applies only to communication sites)
`--owners [owners]`|Comma-separated list of users to set as site owners (applies only to team sites)
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Create modern team site with private group

```sh
spo site add --alias team1 --title Team 1
```

Create modern team site with description and classification

```sh
spo site add --type TeamSite --alias team1 --title Team 1 --description Site of team 1 --classification LBI
```

Create modern team site with public group

```sh
spo site add --type TeamSite --alias team1 --title Team 1 --isPublic
```

Create modern team site using the Dutch language

```sh
spo site add --alias team1 --title Team 1 --lcid 1043
```

Create modern team site with the specified users as owners

```sh
spo site add --alias team1 --title Team 1 --owners "steve@contoso.com, bob@contoso.com"
```

Create communication site using the Topic design

```sh
spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing
```

Create communication site using the Showcase design

```sh
spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --siteDesign Showcase
```

Create communication site using a custom site design

```sh
spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --siteDesignId 99f410fe-dd79-4b9d-8531-f2270c9c621c
```

Create communication site using the Blank design with description and classification

```sh
spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --description Site of the marketing department --classification MBI --siteDesign Blank
```

## More information

- Creating SharePoint Communication Site using REST: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/communication-site-creation-rest](https://docs.microsoft.com/en-us/sharepoint/dev/apis/communication-site-creation-rest)
