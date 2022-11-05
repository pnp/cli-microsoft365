# spo userprofile get

Get SharePoint user profile properties for the specified user

## Usage

```sh
spo userprofile get [options]
```

## Options

`-u, --userName <userName>`
: Account name of the user

--8<-- "docs/cmd/_global.md"

## Remarks

You have to have tenant admin permissions in order to use this command to get profile properties of other users.

## Examples

 Get SharePoint user profile for the specified user

```sh
m365 spo userprofile get --userName 'john.doe@mytenant.onmicrosoft.com'
```

## Response

=== "JSON"

    ```json
    {
      "AccountName": "i:0#.f|membership|johndoe@contoso.onmicrosoft.com",
      "DirectReports": [],
      "DisplayName": "Johndoe",
      "Email": "johndoe@contoso.onmicrosoft.com",
      "ExtendedManagers": [],
      "ExtendedReports": [
        "i:0#.f|membership|johndoe@contoso.onmicrosoft.com"
      ],
      "IsFollowed": false,
      "LatestPost": null,
      "Peers": [],
      "PersonalSiteHostUrl": "https://contoso-my.sharepoint.com:443/",
      "PersonalUrl": "https://contoso-my.sharepoint.com/personal/johndoe_contoso_onmicrosoft_com/",
      "PictureUrl": null,
      "Title": null,
      "UserProfileProperties": [
        {
          "Key": "UserProfile_GUID",
          "Value": "0b4f6da0-97db-456e-993d-80e035057600",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SID",
          "Value": "i:0h.f|membership|100320022ec308a7@live.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "ADGuid",
          "Value": "System.Byte[]",
          "ValueType": "Edm.String"
        },
        {
          "Key": "AccountName",
          "Value": "i:0#.f|membership|johndoe@contoso.onmicrosoft.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "FirstName",
          "Value": "John",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PhoneticFirstName",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "LastName",
          "Value": "Doe",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PhoneticLastName",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "PreferredName",
          "Value": "John Doe",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PhoneticDisplayName",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "WorkPhone",
          "Value": "494594133",
          "ValueType": "Edm.String"
        },
        {
          "Key": "Department",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "Title",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Department",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "Manager",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "AboutMe",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "PersonalSpace",
          "Value": "/personal/johndoe_contosos_onmicrosoft_com/",
          "ValueType": "Edm.String"
        },
        {
          "Key": "PictureURL",
          "Value": "https://contoso-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/78ccf530-bbf0-47e4-aae6-da5f8c6fb142_MThumb.jpg",
          "ValueType": "Edm.String"
        },
        {
          "Key": "UserName",
          "Value": "johndoe@contoso.onmicrosoft.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "QuickLinks",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "WebSite",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "PublicSiteRedirect",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-JobTitle",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-DataSource",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-MemberOf",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Dotted-line",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Peers",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Responsibility",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-SipAddress",
          "Value": "johndoe@contoso.onmicrosoft.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-MySiteUpgrade",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-DontSuggestList",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ProxyAddresses",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-HireDate",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-DisplayOrder",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ClaimID",
          "Value": "johndoe@contoso.onmicrosoft.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ClaimProviderID",
          "Value": "membership",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-LastColleagueAdded",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-OWAUrl",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ResourceSID",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ResourceAccountName",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-MasterAccountName",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-UserPrincipalName",
          "Value": "johndoe@contoso.onmicrosoft.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-O15FirstRunExperience",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PersonalSiteInstantiationState",
          "Value": "2",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-DistinguishedName",
          "Value": "CN=78ccf530-bbf0-47e4-aae6-da5f8c6fb142,OU=0cac6cda-2e04-4a3d-9c16-9c91470d7022,OU=Tenants,OU=MSOnline,DC=SPODS188311,DC=msft,DC=net",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-SourceObjectDN",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-LastKeywordAdded",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ClaimProviderType",
          "Value": "Forms",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-SavedAccountName",
          "Value": "i:0#.f|membership|johndoe@contoso.onmicrosoft.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-SavedSID",
          "Value": "System.Byte[]",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ObjectExists",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PersonalSiteCapabilities",
          "Value": "36",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PersonalSiteFirstCreationTime",
          "Value": "9/12/2022 6:19:29 PM",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PersonalSiteLastCreationTime",
          "Value": "9/12/2022 6:19:29 PM",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PersonalSiteNumberOfRetries",
          "Value": "1",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PersonalSiteFirstCreationError",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-FeedIdentifier",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "WorkEmail",
          "Value": "johndoe@contoso.onmicrosoft.com",
          "ValueType": "Edm.String"
        },
        {
          "Key": "CellPhone",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "Fax",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "HomePhone",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "Office",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Location",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "Assistant",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PastProjects",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Skills",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-School",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Birthday",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-StatusNotes",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Interests",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-HashTags",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-EmailOptin",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PrivacyPeople",
          "Value": "True",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PrivacyActivity",
          "Value": "4095",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PictureTimestamp",
          "Value": "63799442436",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PicturePlaceholderState",
          "Value": "1",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PictureExchangeSyncState",
          "Value": "1",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-MUILanguages",
          "Value": "en-US",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ContentLanguages",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-TimeZone",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-RegionalSettings-FollowWeb",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Locale",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-CalendarType",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-AltCalendarType",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-AdjustHijriDays",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-ShowWeeks",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-WorkDays",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-WorkDayStartHour",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-WorkDayEndHour",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-Time24",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-FirstDayOfWeek",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-FirstWeekOfYear",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-RegionalSettings-Initialized",
          "Value": "True",
          "ValueType": "Edm.String"
        },
        {
          "Key": "OfficeGraphEnabled",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-UserType",
          "Value": "0",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-HideFromAddressLists",
          "Value": "False",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-RecipientTypeDetails",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "DelveFlags",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "PulseMRUPeople",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "msOnline-ObjectId",
          "Value": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-PointPublishingUrl",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-TenantInstanceId",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-SharePointHomeExperienceState",
          "Value": "17301504",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-RefreshToken",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "SPS-MultiGeoFlags",
          "Value": "",
          "ValueType": "Edm.String"
        },
        {
          "Key": "PreferredDataLocation",
          "Value": "",
          "ValueType": "Edm.String"
        }
      ],
      "UserUrl": "https://contoso-my.sharepoint.com:443/Person.aspx?accountname=i%3A0%23%2Ef%7Cmembership%johndoe%40contoso%2Eonmicrosoft%2Ecom"
    }
    ```

=== "Text"

    ```text
    AccountName          : i:0#.f|membership|johndoe@contoso.onmicrosoft.com
    DirectReports        : []
    DisplayName          : John Doe
    Email                : johndoe@contoso.onmicrosoft.com
    ExtendedManagers     : []
    ExtendedReports      : ["i:0#.f|membership|johndoe@contoso.onmicrosoft.com"]
    IsFollowed           : false
    LatestPost           : null
    Peers                : []
    PersonalSiteHostUrl  : https://contoso-my.sharepoint.com:443/
    PersonalUrl          : https://contoso-my.sharepoint.com/personal/johndoe_contoso_onmicrosoft_com/
    PictureUrl           : null
    Title                : null
    UserProfileProperties: [{"Key":"UserProfile_GUID","Value":"0b4f6da0-97db-456e-993d-80e035057600","ValueType":"Edm.String"},{"Key":"SID","Value":"i:0h.f|membership|100320022ec308a7@live.com","ValueType":"Edm.String"},{"Key":"ADGuid","Value":"System.Byte[]","ValueType":"Edm.String"},{"Key":"AccountName","Value":"i:0#.f|membership|johndoe@contoso.onmicrosoft.com","ValueType":"Edm.String"},{"Key":"FirstName","Value":"John","ValueType":"Edm.String"},{"Key":"SPS-PhoneticFirstName","Value":"","ValueType":"Edm.String"},{"Key":"LastName","Value":"Doe","ValueType":"Edm.String"},{"Key":"SPS-PhoneticLastName","Value":"","ValueType":"Edm.String"},{"Key":"PreferredName","Value":"John Doe","ValueType":"Edm.String"},{"Key":"SPS-PhoneticDisplayName","Value":"","ValueType":"Edm.String"},{"Key":"WorkPhone","Value":"494594133","ValueType":"Edm.String"},{"Key":"Department","Value":"","ValueType":"Edm.String"},{"Key":"Title","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Department","Value":"","ValueType":"Edm.String"},{"Key":"Manager","Value":"","ValueType":"Edm.String"},{"Key":"AboutMe","Value":"","ValueType":"Edm.String"},{"Key":"PersonalSpace","Value":"/personal/johndoe_contoso_onmicrosoft_com/","ValueType":"Edm.String"},{"Key":"PictureURL","Value":"https://contoso-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/78ccf530-bbf0-47e4-aae6-da5f8c6fb142_MThumb.jpg","ValueType":"Edm.String"},{"Key":"UserName","Value":"johndoe@contoso.onmicrosoft.com","ValueType":"Edm.String"},{"Key":"QuickLinks","Value":"","ValueType":"Edm.String"},{"Key":"WebSite","Value":"","ValueType":"Edm.String"},{"Key":"PublicSiteRedirect","Value":"","ValueType":"Edm.String"},{"Key":"SPS-JobTitle","Value":"","ValueType":"Edm.String"},{"Key":"SPS-DataSource","Value":"","ValueType":"Edm.String"},{"Key":"SPS-MemberOf","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Dotted-line","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Peers","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Responsibility","Value":"","ValueType":"Edm.String"},{"Key":"SPS-SipAddress","Value":"johndoe@contoso.onmicrosoft.com","ValueType":"Edm.String"},{"Key":"SPS-MySiteUpgrade","Value":"","ValueType":"Edm.String"},{"Key":"SPS-DontSuggestList","Value":"","ValueType":"Edm.String"},{"Key":"SPS-ProxyAddresses","Value":"","ValueType":"Edm.String"},{"Key":"SPS-HireDate","Value":"","ValueType":"Edm.String"},{"Key":"SPS-DisplayOrder","Value":"","ValueType":"Edm.String"},{"Key":"SPS-ClaimID","Value":"johndoe@contoso.onmicrosoft.com","ValueType":"Edm.String"},{"Key":"SPS-ClaimProviderID","Value":"membership","ValueType":"Edm.String"},{"Key":"SPS-LastColleagueAdded","Value":"","ValueType":"Edm.String"},{"Key":"SPS-OWAUrl","Value":"","ValueType":"Edm.String"},{"Key":"SPS-ResourceSID","Value":"","ValueType":"Edm.String"},{"Key":"SPS-ResourceAccountName","Value":"","ValueType":"Edm.String"},{"Key":"SPS-MasterAccountName","Value":"","ValueType":"Edm.String"},{"Key":"SPS-UserPrincipalName","Value":"johndoe@contoso.onmicrosoft.com","ValueType":"Edm.String"},{"Key":"SPS-O15FirstRunExperience","Value":"","ValueType":"Edm.String"},{"Key":"SPS-PersonalSiteInstantiationState","Value":"2","ValueType":"Edm.String"},{"Key":"SPS-DistinguishedName","Value":"CN=78ccf530-bbf0-47e4-aae6-da5f8c6fb142,OU=0cac6cda-2e04-4a3d-9c16-9c91470d7022,OU=Tenants,OU=MSOnline,DC=SPODS188311,DC=msft,DC=net","ValueType":"Edm.String"},{"Key":"SPS-SourceObjectDN","Value":"","ValueType":"Edm.String"},{"Key":"SPS-LastKeywordAdded","Value":"","ValueType":"Edm.String"},{"Key":"SPS-ClaimProviderType","Value":"Forms","ValueType":"Edm.String"},{"Key":"SPS-SavedAccountName","Value":"i:0#.f|membership|johndoe@contoso.onmicrosoft.com","ValueType":"Edm.String"},{"Key":"SPS-SavedSID","Value":"System.Byte[]","ValueType":"Edm.String"},{"Key":"SPS-ObjectExists","Value":"","ValueType":"Edm.String"},{"Key":"SPS-PersonalSiteCapabilities","Value":"36","ValueType":"Edm.String"},{"Key":"SPS-PersonalSiteFirstCreationTime","Value":"9/12/2022 6:19:29 PM","ValueType":"Edm.String"},{"Key":"SPS-PersonalSiteLastCreationTime","Value":"9/12/2022 6:19:29 PM","ValueType":"Edm.String"},{"Key":"SPS-PersonalSiteNumberOfRetries","Value":"1","ValueType":"Edm.String"},{"Key":"SPS-PersonalSiteFirstCreationError","Value":"","ValueType":"Edm.String"},{"Key":"SPS-FeedIdentifier","Value":"","ValueType":"Edm.String"},{"Key":"WorkEmail","Value":"johndoe@contoso.onmicrosoft.com","ValueType":"Edm.String"},{"Key":"CellPhone","Value":"","ValueType":"Edm.String"},{"Key":"Fax","Value":"","ValueType":"Edm.String"},{"Key":"HomePhone","Value":"","ValueType":"Edm.String"},{"Key":"Office","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Location","Value":"","ValueType":"Edm.String"},{"Key":"Assistant","Value":"","ValueType":"Edm.String"},{"Key":"SPS-PastProjects","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Skills","Value":"","ValueType":"Edm.String"},{"Key":"SPS-School","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Birthday","Value":"","ValueType":"Edm.String"},{"Key":"SPS-StatusNotes","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Interests","Value":"","ValueType":"Edm.String"},{"Key":"SPS-HashTags","Value":"","ValueType":"Edm.String"},{"Key":"SPS-EmailOptin","Value":"","ValueType":"Edm.String"},{"Key":"SPS-PrivacyPeople","Value":"True","ValueType":"Edm.String"},{"Key":"SPS-PrivacyActivity","Value":"4095","ValueType":"Edm.String"},{"Key":"SPS-PictureTimestamp","Value":"63799442436","ValueType":"Edm.String"},{"Key":"SPS-PicturePlaceholderState","Value":"1","ValueType":"Edm.String"},{"Key":"SPS-PictureExchangeSyncState","Value":"1","ValueType":"Edm.String"},{"Key":"SPS-MUILanguages","Value":"en-US","ValueType":"Edm.String"},{"Key":"SPS-ContentLanguages","Value":"","ValueType":"Edm.String"},{"Key":"SPS-TimeZone","Value":"","ValueType":"Edm.String"},{"Key":"SPS-RegionalSettings-FollowWeb","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Locale","Value":"","ValueType":"Edm.String"},{"Key":"SPS-CalendarType","Value":"","ValueType":"Edm.String"},{"Key":"SPS-AltCalendarType","Value":"","ValueType":"Edm.String"},{"Key":"SPS-AdjustHijriDays","Value":"","ValueType":"Edm.String"},{"Key":"SPS-ShowWeeks","Value":"","ValueType":"Edm.String"},{"Key":"SPS-WorkDays","Value":"","ValueType":"Edm.String"},{"Key":"SPS-WorkDayStartHour","Value":"","ValueType":"Edm.String"},{"Key":"SPS-WorkDayEndHour","Value":"","ValueType":"Edm.String"},{"Key":"SPS-Time24","Value":"","ValueType":"Edm.String"},{"Key":"SPS-FirstDayOfWeek","Value":"","ValueType":"Edm.String"},{"Key":"SPS-FirstWeekOfYear","Value":"","ValueType":"Edm.String"},{"Key":"SPS-RegionalSettings-Initialized","Value":"True","ValueType":"Edm.String"},{"Key":"OfficeGraphEnabled","Value":"","ValueType":"Edm.String"},{"Key":"SPS-UserType","Value":"0","ValueType":"Edm.String"},{"Key":"SPS-HideFromAddressLists","Value":"False","ValueType":"Edm.String"},{"Key":"SPS-RecipientTypeDetails","Value":"","ValueType":"Edm.String"},{"Key":"DelveFlags","Value":"","ValueType":"Edm.String"},{"Key":"PulseMRUPeople","Value":"","ValueType":"Edm.String"},{"Key":"msOnline-ObjectId","Value":"78ccf530-bbf0-47e4-aae6-da5f8c6fb142","ValueType":"Edm.String"},{"Key":"SPS-PointPublishingUrl","Value":"","ValueType":"Edm.String"},{"Key":"SPS-TenantInstanceId","Value":"","ValueType":"Edm.String"},{"Key":"SPS-SharePointHomeExperienceState","Value":"17301504","ValueType":"Edm.String"},{"Key":"SPS-RefreshToken","Value":"","ValueType":"Edm.String"},{"Key":"SPS-MultiGeoFlags","Value":"","ValueType":"Edm.String"},{"Key":"PreferredDataLocation","Value":"","ValueType":"Edm.String"}]
    UserUrl              : https://contoso-my.sharepoint.com:443/Person.aspx?accountname=i%3A0%23%2Ef%7Cmembership%7Cjohndoe%40contoso%2Eonmicrosoft%2Ecom
    ```

=== "CSV"

    ```csv
    AccountName,DirectReports,DisplayName,Email,ExtendedManagers,ExtendedReports,IsFollowed,LatestPost,Peers,PersonalSiteHostUrl,PersonalUrl,PictureUrl,Title,UserProfileProperties,UserUrl
    i:0#.f|membership|johndoe@contoso.onmicrosoft.com,[],John Doe,johndoe@contoso.onmicrosoft.com,[],"[""i:0#.f|membership|johndoe@contoso.onmicrosoft.com""]",,,[],https://contoso-my.sharepoint.com:443/,https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com/,,,"[{""Key"":""UserProfile_GUID"",""Value"":""0b4f6da0-97db-456e-993d-80e035057600"",""ValueType"":""Edm.String""},{""Key"":""SID"",""Value"":""i:0h.f|membership|100320022ec308a7@live.com"",""ValueType"":""Edm.String""},{""Key"":""ADGuid"",""Value"":""System.Byte[]"",""ValueType"":""Edm.String""},{""Key"":""AccountName"",""Value"":""i:0#.f|membership|johndoe@contoso.onmicrosoft.com"",""ValueType"":""Edm.String""},{""Key"":""FirstName"",""Value"":""John"",""ValueType"":""Edm.String""},{""Key"":""SPS-PhoneticFirstName"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""LastName"",""Value"":""Doe"",""ValueType"":""Edm.String""},{""Key"":""SPS-PhoneticLastName"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""PreferredName"",""Value"":""John Doe"",""ValueType"":""Edm.String""},{""Key"":""SPS-PhoneticDisplayName"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""WorkPhone"",""Value"":""494594133"",""ValueType"":""Edm.String""},{""Key"":""Department"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""Title"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Department"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""Manager"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""AboutMe"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""PersonalSpace"",""Value"":""/personal/john_contoso_onmicrosoft_com/"",""ValueType"":""Edm.String""},{""Key"":""PictureURL"",""Value"":""https://contoso-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/78ccf530-bbf0-47e4-aae6-da5f8c6fb142_MThumb.jpg"",""ValueType"":""Edm.String""},{""Key"":""UserName"",""Value"":""johndoe@contoso.onmicrosoft.com"",""ValueType"":""Edm.String""},{""Key"":""QuickLinks"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""WebSite"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""PublicSiteRedirect"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-JobTitle"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-DataSource"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-MemberOf"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Dotted-line"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Peers"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Responsibility"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-SipAddress"",""Value"":""johndoe@contoso.onmicrosoft.com"",""ValueType"":""Edm.String""},{""Key"":""SPS-MySiteUpgrade"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-DontSuggestList"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-ProxyAddresses"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-HireDate"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-DisplayOrder"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-ClaimID"",""Value"":""johndoe@contoso.onmicrosoft.com"",""ValueType"":""Edm.String""},{""Key"":""SPS-ClaimProviderID"",""Value"":""membership"",""ValueType"":""Edm.String""},{""Key"":""SPS-LastColleagueAdded"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-OWAUrl"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-ResourceSID"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-ResourceAccountName"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-MasterAccountName"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-UserPrincipalName"",""Value"":""johndoe@contoso.onmicrosoft.com"",""ValueType"":""Edm.String""},{""Key"":""SPS-O15FirstRunExperience"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-PersonalSiteInstantiationState"",""Value"":""2"",""ValueType"":""Edm.String""},{""Key"":""SPS-DistinguishedName"",""Value"":""CN=78ccf530-bbf0-47e4-aae6-da5f8c6fb142,OU=0cac6cda-2e04-4a3d-9c16-9c91470d7022,OU=Tenants,OU=MSOnline,DC=SPODS188311,DC=msft,DC=net"",""ValueType"":""Edm.String""},{""Key"":""SPS-SourceObjectDN"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-LastKeywordAdded"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-ClaimProviderType"",""Value"":""Forms"",""ValueType"":""Edm.String""},{""Key"":""SPS-SavedAccountName"",""Value"":""i:0#.f|membership|johndoe@contoso.onmicrosoft.com"",""ValueType"":""Edm.String""},{""Key"":""SPS-SavedSID"",""Value"":""System.Byte[]"",""ValueType"":""Edm.String""},{""Key"":""SPS-ObjectExists"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-PersonalSiteCapabilities"",""Value"":""36"",""ValueType"":""Edm.String""},{""Key"":""SPS-PersonalSiteFirstCreationTime"",""Value"":""9/12/2022 6:19:29 PM"",""ValueType"":""Edm.String""},{""Key"":""SPS-PersonalSiteLastCreationTime"",""Value"":""9/12/2022 6:19:29 PM"",""ValueType"":""Edm.String""},{""Key"":""SPS-PersonalSiteNumberOfRetries"",""Value"":""1"",""ValueType"":""Edm.String""},{""Key"":""SPS-PersonalSiteFirstCreationError"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-FeedIdentifier"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""WorkEmail"",""Value"":""johndoe@contoso.onmicrosoft.com"",""ValueType"":""Edm.String""},{""Key"":""CellPhone"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""Fax"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""HomePhone"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""Office"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Location"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""Assistant"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-PastProjects"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Skills"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-School"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Birthday"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-StatusNotes"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Interests"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-HashTags"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-EmailOptin"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-PrivacyPeople"",""Value"":""True"",""ValueType"":""Edm.String""},{""Key"":""SPS-PrivacyActivity"",""Value"":""4095"",""ValueType"":""Edm.String""},{""Key"":""SPS-PictureTimestamp"",""Value"":""63799442436"",""ValueType"":""Edm.String""},{""Key"":""SPS-PicturePlaceholderState"",""Value"":""1"",""ValueType"":""Edm.String""},{""Key"":""SPS-PictureExchangeSyncState"",""Value"":""1"",""ValueType"":""Edm.String""},{""Key"":""SPS-MUILanguages"",""Value"":""en-US"",""ValueType"":""Edm.String""},{""Key"":""SPS-ContentLanguages"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-TimeZone"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-RegionalSettings-FollowWeb"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Locale"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-CalendarType"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-AltCalendarType"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-AdjustHijriDays"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-ShowWeeks"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-WorkDays"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-WorkDayStartHour"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-WorkDayEndHour"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-Time24"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-FirstDayOfWeek"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-FirstWeekOfYear"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-RegionalSettings-Initialized"",""Value"":""True"",""ValueType"":""Edm.String""},{""Key"":""OfficeGraphEnabled"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-UserType"",""Value"":""0"",""ValueType"":""Edm.String""},{""Key"":""SPS-HideFromAddressLists"",""Value"":""False"",""ValueType"":""Edm.String""},{""Key"":""SPS-RecipientTypeDetails"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""DelveFlags"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""PulseMRUPeople"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""msOnline-ObjectId"",""Value"":""78ccf530-bbf0-47e4-aae6-da5f8c6fb142"",""ValueType"":""Edm.String""},{""Key"":""SPS-PointPublishingUrl"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-TenantInstanceId"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-SharePointHomeExperienceState"",""Value"":""17301504"",""ValueType"":""Edm.String""},{""Key"":""SPS-RefreshToken"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""SPS-MultiGeoFlags"",""Value"":"""",""ValueType"":""Edm.String""},{""Key"":""PreferredDataLocation"",""Value"":"""",""ValueType"":""Edm.String""}]",https://contoso-my.sharepoint.com:443/Person.aspx?accountname=i%3A0%23%2Ef%7Cmembership%7Cjohn%40contoso%2Eonmicrosoft%2Ecom
    ```


