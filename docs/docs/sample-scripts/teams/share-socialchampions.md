# Share social champions to Teams

Author: [Albert-Jan Schot](https://www.cloudappie.nl/recognize-contributions-clim365) Inspired by: [Emily Mancini](https://docs.microsoft.com/en-us/microsoft-365/community/identifying-your-sharepoint-champions)

Retrieves activities for SharePoint Online, Teams and Yammer and shares the top 3 contributors for each category as an adaptive card to the specified webhook url.

```powershell tab="PowerShell"
$m365Status = m365 status

if ($m365Status -eq "Logged Out") {
  # Connection to Microsoft 365
  m365 login
}

$webhookUrl = "<PUTYOURURLHERE>"

# Send top 3 for SharePoint based on file actions.
$activityUsers = m365 spo report activityuserdetail --period D7 --output json --query 'reverse(sort_by(@, &\"Viewed Or Edited File Count\")) | [0:3].\"User Principal Name\"' | ConvertFrom-Json
$title = "SharePoint Weekly Social Champions"
$card = '{ \"type\": \"AdaptiveCard\", \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\", \"version\": \"1.2\", \"body\": [  {  \"type\": \"TextBlock\",  \"text\": \"'+$($title)+'\",  \"wrap\": true,  \"size\": \"Medium\",  \"weight\": \"Bolder\",  \"color\": \"Attention\"  },  {  \"type\": \"TextBlock\",  \"wrap\": true,  \"text\": \"Week '+$(get-date -UFormat %V)+'\",  \"fontType\": \"Default\",  \"size\": \"Small\",  \"weight\": \"Lighter\",  \"isSubtle\": true  },  {  \"type\": \"FactSet\",  \"facts\": [   {   \"title\": \"First place\",   \"value\": \"'+$($activityUsers[0])+'\"   },   {   \"title\": \"Second place\",   \"value\": \"'+$($activityUsers[1])+'\"   },   {   \"title\": \"Third place\",   \"value\": \"'+$($activityUsers[2])+'\"   }  ]  } ] }'
m365 adaptivecard send --url $webhookUrl --card $card

# Send top 3 for Teams based on chat messages
$activityUsers = m365 teams report useractivityuserdetail --period D7 --output json --query 'reverse(sort_by(@, &\"Team Chat Message Count\")) | [0:3].\"User Principal Name\"' | ConvertFrom-Json
$title = "Teams Weekly Social Champions"
$card = '{ \"type\": \"AdaptiveCard\", \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\", \"version\": \"1.2\", \"body\": [  {  \"type\": \"TextBlock\",  \"text\": \"'+$($title)+'\",  \"wrap\": true,  \"size\": \"Medium\",  \"weight\": \"Bolder\",  \"color\": \"Attention\"  },  {  \"type\": \"TextBlock\",  \"wrap\": true,  \"text\": \"Week '+$(get-date -UFormat %V)+'\",  \"fontType\": \"Default\",  \"size\": \"Small\",  \"weight\": \"Lighter\",  \"isSubtle\": true  },  {  \"type\": \"FactSet\",  \"facts\": [   {   \"title\": \"First place\",   \"value\": \"'+$($activityUsers[0])+'\"   },   {   \"title\": \"Second place\",   \"value\": \"'+$($activityUsers[1])+'\"   },   {   \"title\": \"Third place\",   \"value\": \"'+$($activityUsers[2])+'\"   }  ]  } ] }'
m365 adaptivecard send --url $webhookUrl --card $card

# Send top 3 for Yammer based on posts
$activityUsers = m365 yammer report activityuserdetail --period D7 --output json --query 'reverse(sort_by(@, &\"Posted Count\")) | [0:3].\"User Principal Name\"' | ConvertFrom-Json
$title = "Yammer Weekly Social Champions"
$card = '{ \"type\": \"AdaptiveCard\", \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\", \"version\": \"1.2\", \"body\": [  {  \"type\": \"TextBlock\",  \"text\": \"'+$($title)+'\",  \"wrap\": true,  \"size\": \"Medium\",  \"weight\": \"Bolder\",  \"color\": \"Attention\"  },  {  \"type\": \"TextBlock\",  \"wrap\": true,  \"text\": \"Week '+$(get-date -UFormat %V)+'\",  \"fontType\": \"Default\",  \"size\": \"Small\",  \"weight\": \"Lighter\",  \"isSubtle\": true  },  {  \"type\": \"FactSet\",  \"facts\": [   {   \"title\": \"First place\",   \"value\": \"'+$($activityUsers[0])+'\"   },   {   \"title\": \"Second place\",   \"value\": \"'+$($activityUsers[1])+'\"   },   {   \"title\": \"Third place\",   \"value\": \"'+$($activityUsers[2])+'\"   }  ]  } ] }'
m365 adaptivecard send --url $webhookUrl --card $card
```

Keywords:

- Microsoft Teams
- SharePoint Online
- Yammer
- Adoption
