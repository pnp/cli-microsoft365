# spo list add

Creates list in the specified site

## Usage

```sh
m365 spo list add [options]
```

## Options

`-h, --help`
: output usage information

`-t
: --title <title>`|Title of the list to add

`--baseTemplate <baseTemplate>`
: The list definition type on which the list is based. Allowed values `Announcements,Contacts,CustomGrid,DataSources,DiscussionBoard,DocumentLibrary,Events,GanttTasks,GenericList,IssuesTracking,Links,NoCodeWorkflows,PictureLibrary,Survey,Tasks,WebPageLibrary,WorkflowHistory,WorkflowProcess,XmlForm`. Default GenericList

`-u
: --webUrl <webUrl>`|URL of the site where the list should be added

`--description [description]`
: The description for the list

`--templateFeatureId [templateFeatureId]`
: The globally unique identifier (GUID) of a template feature that is associated with the list

`--schemaXml [schemaXml]`
: The schema in Collaborative Application Markup Language (CAML) schemas that defines the list

`--allowDeletion [allowDeletion]`
: Boolean value specifying whether the list can be deleted. Valid values are `true,false`

`--allowEveryoneViewItems [allowEveryoneViewItems]`
: Boolean value specifying whether everyone can view documents in the documentlibrary or attachments to items in the list. Valid values are `true,false`

`--allowMultiResponses [allowMultiResponses]`
: Boolean value specifying whether users are allowed to give multiple responses to the survey. Valid values are `true,false`

`--contentTypesEnabled [contentTypesEnabled]`
: Boolean value specifying whether content types are enabled for the list. Valid values are `true,false`

`--crawlNonDefaultViews [crawlNonDefaultViews]`
: Boolean value specifying whether to crawl non default views. Valid values are `true,false`

`--defaultContentApprovalWorkflowId [defaultContentApprovalWorkflowId]`
: Value that specifies the default workflow identifier for content approval onthe list (GUID)

`--defaultDisplayFormUrl [defaultDisplayFormUrl]`
: Value that specifies the location of the default display form for the list

`--defaultEditFormUrl [defaultEditFormUrl]`
: Value that specifies the URL of the edit form to use for list items in the list

`--direction [direction]`
: Value that specifies the reading order of the list. Valid values are `NONE,LTR,RTL`

`--disableGridEditing [disableGridEditing]`
: Property for assigning or retrieving grid editing on the list. Valid values are `true,false`

`--draftVersionVisibility [draftVersionVisibility]`
: Value that specifies the minimum permission required to view minor versions and drafts within the list. Allowed values `Reader,Author,Approver`. Default Reader

`--emailAlias [emailAlias]`
: If e-mail notification is enabled, gets or sets the e-mail address to use tonotify to the owner of an item when an assignment has changed or the item has been updated.

`--enableAssignToEmail [enableAssignToEmail]`
: Boolean value specifying whether e-mail notification is enabled for the list. Valid values are `true,false`

`--enableAttachments [enableAttachments]`
: Boolean value that specifies whether attachments can be added to items in the list. Valid values are `true,false`

`--enableDeployWithDependentList [enableDeployWithDependentList]`
: Boolean value that specifies whether the list can be deployed with a dependent list. Valid values are `true,false`

`--enableFolderCreation [enableFolderCreation]`
: Boolean value that specifies whether folders can be created for the list. Valid values are `true,false`

`--enableMinorVersions [enableMinorVersions]`
: Boolean value that specifies whether minor versions are enabled when versioning is enabled for the document library. Valid values are `true,false`

`--enableModeration [enableModeration]`
: Boolean value that specifies whether Content Approval is enabled for the list. Valid values are `true,false`

`--enablePeopleSelector [enablePeopleSelector]`
: Enable user selector on event list. Valid values are `true,false`

`--enableResourceSelector [enableResourceSelector]`
: Enables resource selector on an event list. Valid values are `true,false`

`--enableSchemaCaching [enableSchemaCaching]`
: Boolean value specifying whether schema caching is enabled for the list. Valid values are `true,false`

`--enableSyndication [enableSyndication]`
: Boolean value that specifies whether RSS syndication is enabled for the list. Valid values are `true,false`

`--enableThrottling [enableThrottling]`
: Indicates whether throttling for this list is enabled or not. Valid values are `true,false`

`--enableVersioning [enableVersioning]`
: Boolean value that specifies whether versioning is enabled for the document library. Valid values are `true,false`

`--enforceDataValidation [enforceDataValidation]`
: Value that indicates whether certain field properties are enforced when an item is added or updated. Valid values are `true,false`

`--excludeFromOfflineClient [excludeFromOfflineClient]`
: Value that indicates whether the list should be downloaded to the client during offline synchronization. Valid values are `true,false`

`--fetchPropertyBagForListView [fetchPropertyBagForListView]`
: Specifies whether property bag information, as part of the list schema JSON,is retrieved when the list is being rendered on the client. Valid values are `true,false`

`--followable [followable]`
: Can a list be followed in an activity feed?. Valid values are `true,false`

`--forceCheckout [forceCheckout]`
: Boolean value that specifies whether forced checkout is enabled for the document library. Valid values are `true,false`

`--forceDefaultContentType [forceDefaultContentType]`
: Specifies whether we want to return the default Document root content type. Valid values are `true,false`

`--hidden [hidden]`
: Boolean value that specifies whether the list is hidden. Valid values are `true,false`

`--includedInMyFilesScope [includedInMyFilesScope]`
: Specifies whether this list is accessible to an app principal that has been granted an OAuth scope that contains the string “myfiles” by a case-insensitive comparison when the current user is a site collection administrator of the personal site that contains the list

`--irmEnabled [irmEnabled]`
: Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) is enabled for the list

`--irmExpire [irmExpire]`
: Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) expiration is enabled for the list

`--irmReject [irmReject]`
: Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) rejection is enabled for the list

`--isApplicationList [isApplicationList]`
: Indicates whether this list should be treated as a top level navigation object or not

`--listExperienceOptions [listExperienceOptions]`
: Gets or sets the list experience for the list. Allowed values Auto,NewExperience,ClassicExperience. Default Auto

`--majorVersionLimit [majorVersionLimit]`
: Gets or sets the maximum number of major versions allowed for an item in a document library that uses version control with major versions only.

`--majorWithMinorVersionsLimit [majorWithMinorVersionsLimit]`
: Gets or sets the maximum number of major versions that are allowed for an item in a document library that uses version control with both major and minor versions.

`--multipleDataList [multipleDataList]`
: Gets or sets a Boolean value that specifies whether the list in a Meeting Workspace sitecontains data for multiple meeting instances within the site

`--navigateForFormsPages [navigateForFormsPages]`
: Indicates whether to navigate for forms pages or use a modal dialog

`--needUpdateSiteClientTag [needUpdateSiteClientTag]`
: A boolean value that determines whether to editing documents in this list should increment the ClientTag for the site. The tag is used to allow clients to cache JS/CSS/resources that are retrieved from the Content DB, including custom CSR templates.

`--noCrawl [noCrawl]`
: Gets or sets a Boolean value specifying whether crawling is enabled for the list

`--onQuickLaunch [onQuickLaunch]`
: Gets or sets a Boolean value that specifies whether the list appears on the Quick Launcharea of the home page

`--ordered [ordered]`
: Gets or sets a Boolean value that specifies whether the option to allow users to reorderitems in the list is available on the Edit View page for the list

`--parserDisabled [parserDisabled]`
: Gets or sets a Boolean value that specifies whether the parser should be disabled

`--readOnlyUI [readOnlyUI]`
: A boolean value that indicates whether the UI for this list should be presented in a read-only fashion. This will not affect security nor will it actually prevent changes to the list from occurring - it only affects the way the UI is displayed

`--readSecurity [readSecurity]`
: Gets or sets the Read security setting for the list. Valid values are 1 (All users have Read access to all items)|2 (Users have Read access only to items that they create)

`--requestAccessEnabled [requestAccessEnabled]`
: Gets or sets a Boolean value that specifies whether the option to allow users to requestaccess to the list is available

`--restrictUserUpdates [restrictUserUpdates]`
: A boolean value that indicates whether the this list is a restricted one or not The value can't be changed if there are existing items in the list

`--sendToLocationName [sendToLocationName]`
: Gets or sets a file name to use when copying an item in the list to another document library.

`--sendToLocationUrl [sendToLocationUrl]`
: Gets or sets a URL to use when copying an item in the list to another document library

`--showUser [showUser]`
: Gets or sets a Boolean value that specifies whether names of users are shown in the results of the survey

`--useFormsForDisplay [useFormsForDisplay]`
: Indicates whether forms should be considered for display context or not

`--validationFormula [validationFormula]`
: Gets or sets a formula that is evaluated each time that a list item is added or updated.

`--validationMessage [validationMessage]`
: Gets or sets the message that is displayed when validation fails for a list item.

`--writeSecurity [writeSecurity]`
: Gets or sets the Write security setting for the list. Valid values are 1 (All users can modify all items)|2 (Users can modify only items that they create)|4 (Users cannot modify any list item)

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Add a list with title _Announcements_ and baseTemplate _Announcements_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list add --title Announcements --baseTemplate Announcements --webUrl https://contoso.sharepoint.com/sites/project-x
```

Add a list with title _Announcements_, baseTemplate _Announcements_ in site _https://contoso.sharepoint.com/sites/project-x_ using a custom XML schema

```sh
m365 spo list add --title Announcements --baseTemplate Announcements --webUrl https://contoso.sharepoint.com/sites/project-x --schemaXml '<List xmlns:ows="Microsoft SharePoint" Title="List1" FolderCreation="FALSE" Direction="$Resources:Direction;" Url="Lists/List1" BaseType="0" xmlns="http://schemas.microsoft.com/sharepoint/"><MetaData><ContentTypes><ContentTypeRef ID="0x01"><Folder TargetName="Item" /></ContentTypeRef><ContentTypeRef ID="0x0120" /></ContentTypes><Fields><Field ID="{fa564e0f-0c70-4ab9-b863-0177e6ddd247}" Type="Text" Name="Title" DisplayName="$Resources:core,Title;" Required="TRUE" SourceID="http://schemas.microsoft.com/sharepoint/v3" StaticName="Title" MaxLength="255" /></Fields><Views><View BaseViewID="0" Type="HTML" MobileView="TRUE" TabularView="FALSE"><Toolbar Type="Standard" /><XslLink Default="TRUE">main.xsl</XslLink><RowLimit Paged="TRUE">30</RowLimit><ViewFields><FieldRef Name="LinkTitleNoMenu"></FieldRef></ViewFields><Query><OrderBy><FieldRef Name="Modified" Ascending="FALSE"></FieldRef></OrderBy></Query><ParameterBindings><ParameterBinding Name="AddNewAnnouncement" Location="Resource(wss,addnewitem)" /><ParameterBinding Name="NoAnnouncements" Location="Resource(wss,noXinviewofY_LIST)" /><ParameterBinding Name="NoAnnouncementsHowTo" Location="Resource(wss,noXinviewofY_ONET_HOME)" /></ParameterBindings></View><View BaseViewID="1" Type="HTML" WebPartZoneID="Main" DisplayName="$Resources:core,objectiv_schema_mwsidcamlidC24;" DefaultView="TRUE" MobileView="TRUE" MobileDefaultView="TRUE" SetupPath="pages\viewpage.aspx" ImageUrl="/_layouts/15/images/generic.png?rev=23" Url="AllItems.aspx"><Toolbar Type="Standard" /><XslLink Default="TRUE">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged="TRUE">30</RowLimit><ViewFields><FieldRef Name="LinkTitle"></FieldRef></ViewFields><Query><OrderBy><FieldRef Name="ID"></FieldRef></OrderBy></Query><ParameterBindings><ParameterBinding Name="NoAnnouncements" Location="Resource(wss,noXinviewofY_LIST)" /><ParameterBinding Name="NoAnnouncementsHowTo" Location="Resource(wss,noXinviewofY_DEFAULT)" /></ParameterBindings></View></Views><Forms><Form Type="DisplayForm" Url="DispForm.aspx" SetupPath="pages\form.aspx" WebPartZoneID="Main" /><Form Type="EditForm" Url="EditForm.aspx" SetupPath="pages\form.aspx" WebPartZoneID="Main" /><Form Type="NewForm" Url="NewForm.aspx" SetupPath="pages\form.aspx" WebPartZoneID="Main" /></Forms></MetaData></List>'
```

Add a list with title _Announcements_, baseTemplate _Announcements_ in site _https://contoso.sharepoint.com/sites/project-x_ with content types and versioning enabled and major version limit set to _50_

```sh
m365 spo list add --webUrl https://contoso.sharepoint.com/sites/project-x --title Announcements --baseTemplate Announcements --contentTypesEnabled true --enableVersioning true --majorVersionLimit 50
```

## More information

- SPList Class Members information: [https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.list_members.aspx](https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.list_members.aspx)
- ListTemplateType enum information: [https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listtemplatetype.aspx](https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listtemplatetype.aspx)
- DraftVersionVisibilityType enum information: [https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.draftvisibilitytype.aspx](https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.draftvisibilitytype.aspx)
- ListExperience enum information: [https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listexperience.aspx](https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listexperience.aspx)