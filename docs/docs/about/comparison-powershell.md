# Comparison to SharePoint and Microsoft 365 PowerShell

Following table lists the different CLI for Microsoft 365 commands and how they map to PowerShell cmdlets for SharePoint and Microsoft 365.

PowerShell Cmdlet|Source|CLI for Microsoft 365 command
-----------------|------|----------------------
Add-SPOGeoAdministrator|Microsoft.Online.SharePoint.PowerShell|
Add-SPOHubSiteAssociation|Microsoft.Online.SharePoint.PowerShell|[spo hubsite connect](../cmd/spo/hubsite/hubsite-connect.md)
Add-SPOOrgAssetsLibrary|Microsoft.Online.SharePoint.PowerShell|
Add-SPOSiteCollectionAppCatalog|Microsoft.Online.SharePoint.PowerShell|[spo site appcatalog add](../cmd/spo/site/site-appcatalog-add.md)
Add-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign add](../cmd/spo/sitedesign/sitedesign-add.md)
Add-SPOSiteDesignTask|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
Add-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|[spo sitescript add](../cmd/spo/sitescript/sitescript-add.md)
Add-SPOTenantCdnOrigin|Microsoft.Online.SharePoint.PowerShell|[spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md)
Add-SPOTheme|Microsoft.Online.SharePoint.PowerShell|[spo theme set](../cmd/spo/theme/theme-set.md)
Add-SPOHubToHubAssociation|Microsoft.Online.SharePoint.PowerShell|
Add-SPOSiteScriptPackage|Microsoft.Online.SharePoint.PowerShell|
Add-SPOUser|Microsoft.Online.SharePoint.PowerShell|
Approve-SPOTenantServicePrincipalPermissionGrant|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal grant add](../cmd/spo/serviceprincipal/serviceprincipal-grant-add.md)
Approve-SPOTenantServicePrincipalPermissionRequest|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal permissionrequest approve](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md)
Connect-SPOService|Microsoft.Online.SharePoint.PowerShell|[spo login](../cmd/login.md)
ConvertTo-SPOMigrationEncryptedPackage|Microsoft.Online.SharePoint.PowerShell|
ConvertTo-SPOMigrationTargetedPackage|Microsoft.Online.SharePoint.PowerShell|
Deny-SPOTenantServicePrincipalPermissionRequest|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal permissionrequest deny](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md)
Disable-SPOTenantServicePrincipal|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Disconnect-SPOService|Microsoft.Online.SharePoint.PowerShell|[spo logout](../cmd/logout.md)
Enable-SPOCommSite|Microsoft.Online.SharePoint.PowerShell|[spo site commsite enable](../cmd/spo/site/site-commsite-enable.md)
Enable-SPOTenantServicePrincipal|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Export-SPOQueryLogs|Microsoft.Online.SharePoint.PowerShell|
Export-SPOUserInfo|Microsoft.Online.SharePoint.PowerShell|
Export-SPOUserProfile|Microsoft.Online.SharePoint.PowerShell|
Get-IsCommSite|Microsoft.Online.SharePoint.PowerShell|
Get-SPOAppErrors|Microsoft.Online.SharePoint.PowerShell|
Get-SPOAppInfo|Microsoft.Online.SharePoint.PowerShell|
Get-SPOBrowserIdleSignOut|Microsoft.Online.SharePoint.PowerShell|
Get-SPOBuiltDesignPackageVisibility|Microsoft.Online.SharePoint.PowerShell|
Get-SPOBuiltInDesignPackageVisibility|Microsoft.Online.SharePoint.PowerShell|
Get-SPOCrossGeoMovedUsers|Microsoft.Online.SharePoint.PowerShell|
Get-SPOCrossGeoMoveReport|Microsoft.Online.SharePoint.PowerShell|
Get-SPOCrossGeoUsers|Microsoft.Online.SharePoint.PowerShell|
Get-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Get-SPODeletedSite|Microsoft.Online.SharePoint.PowerShell|[spo site list](../cmd/spo/site/site-list.md)
Get-SPOExternalUser|Microsoft.Online.SharePoint.PowerShell|[spo externaluser list](../cmd/spo/externaluser/externaluser-list.md)
Get-SPOGeoAdministrator|Microsoft.Online.SharePoint.PowerShell|
Get-SPOGeoMoveCrossCompatibilityStatus|Microsoft.Online.SharePoint.PowerShell|
Get-SPOGeoStorageQuota|Microsoft.Online.SharePoint.PowerShell|
Get-SPOHideDefaultThemes|Microsoft.Online.SharePoint.PowerShell|[spo hidedefaultthemes get](../cmd/spo/hidedefaultthemes/hidedefaultthemes-get.md)
Get-SPOHomeSite|Microsoft.Online.SharePoint.PowerShell|[spo homesite get](../cmd/spo/homesite/homesite-get.md)
Get-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite get](../cmd/spo/hubsite/hubsite-get.md), [spo hubsite list](../cmd/spo/hubsite/hubsite-list.md)
Get-SPOKnowledgeHubSite|Microsoft.Online.SharePoint.PowerShell|
Get-SPOMigrationJobProgress|Microsoft.Online.SharePoint.PowerShell|
Get-SPOMigrationJobStatus|Microsoft.Online.SharePoint.PowerShell|
Get-SPOMultiGeoCompanyAllowedDataLocation|Microsoft.Online.SharePoint.PowerShell|
Get-SPOMultiGeoExperience|Microsoft.Online.SharePoint.PowerShell|
Get-SPOOrgAssetsLibrary|Microsoft.Online.SharePoint.PowerShell|[spo orgassetslibrary list](../cmd/spo/orgassetslibrary/orgassetslibrary-list.md)
Get-SPOOrgNewsSite|Microsoft.Online.SharePoint.PowerShell|
Get-SPOPublicCdnOrigins|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSite|Microsoft.Online.SharePoint.PowerShell|[spo site classic list](../cmd/spo/site/site-classic-list.md)
Get-SPOSiteCollectionAppCatalogs|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteContentMoveState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteDataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign get](../cmd/spo/sitedesign/sitedesign-get.md), [spo sitedesign list](../cmd/spo/sitedesign/sitedesign-list.md)
Get-SPOSiteDesignRights|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign rights list](../cmd/spo/sitedesign/sitedesign-rights-list.md)
Get-SPOSiteDesignRun|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign run list](../cmd/spo/sitedesign/sitedesign-run-list.md)
Get-SPOSiteDesignRunStatus|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign run status get](../cmd/spo/sitedesign/sitedesign-run-status-get.md)
Get-SPOSiteDesignTask|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign task get](../cmd/spo/sitedesign/sitedesign-task-get.md), [spo sitedesign task list](../cmd/spo/sitedesign/sitedesign-task-list.md)
Get-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteRenameState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|[spo sitescript get](../cmd/spo/sitescript/sitescript-get.md), [spo sitescript list](../cmd/spo/sitescript/sitescript-list.md)
Get-SPOSiteScriptFromList|Microsoft.Online.SharePoint.PowerShell|[spo list sitescript get](../cmd/spo/list/list-sitescript-get.md)
Get-SPOSiteScriptFromWeb|Microsoft.Online.SharePoint.PowerShell|
Get-SPOSiteUserInvitations|Microsoft.Online.SharePoint.PowerShell|
Get-SPOStorageEntity|Microsoft.Online.SharePoint.PowerShell|[spo storageentity get](../cmd/spo/storageentity/storageentity-get.md), [spo storageentity list](../cmd/spo/storageentity/storageentity-list.md)
Get-SPOStructuralNavigationCacheSiteState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOStructuralNavigationCacheWebState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenant|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantCdnEnabled|Microsoft.Online.SharePoint.PowerShell|[spo cdn get](../cmd/spo/cdn/cdn-get.md)
Get-SPOTenantCdnOrigins|Microsoft.Online.SharePoint.PowerShell|[spo cdn origin list](../cmd/spo/cdn/cdn-origin-list.md)
Get-SPOTenantCdnPolicies|Microsoft.Online.SharePoint.PowerShell|[spo cdn policy list](../cmd/spo/cdn/cdn-policy-list.md)
Get-SPOTenantContentTypeReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantLogEntry|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantLogLastAvailableTimeInUtc|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantOrgRelation|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantOrgRelationByPartner|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantOrgRelationByScenario|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantServicePrincipalPermissionGrants|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal grant list](../cmd/spo/serviceprincipal/serviceprincipal-grant-list.md)
Get-SPOTenantServicePrincipalPermissionRequests|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal permissionrequest list](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-list.md)
Get-SPOTenantSyncClientRestriction|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTenantTaxonomyReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Get-SPOTheme|Microsoft.Online.SharePoint.PowerShell|[spo theme list](../cmd/spo/theme/theme-list.md)
Get-SPOUnifiedGroup|Microsoft.Online.SharePoint.PowerShell|
Get-SPOUnifiedGroupMoveState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOUser|Microsoft.Online.SharePoint.PowerShell|[spo user get](../cmd/spo/user/user-get.md), [spo user list](../cmd/spo/user/user-list.md)
Get-SPOUserAndContentMoveState|Microsoft.Online.SharePoint.PowerShell|
Get-SPOUserOneDriveLocation|Microsoft.Online.SharePoint.PowerShell|
Get-SPOWebTemplate|Microsoft.Online.SharePoint.PowerShell|
Grant-SPOHubSiteRights|Microsoft.Online.SharePoint.PowerShell|[spo hubsite rights grant](../cmd/spo/hubsite/hubsite-rights-grant.md)
Grant-SPOSiteDesignRights|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign rights grant](../cmd/spo/sitedesign/sitedesign-rights-grant.md)
Invoke-SPOMigrationEncryptUploadSubmit|Microsoft.Online.SharePoint.PowerShell|
Invoke-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
Invoke-SPOSiteSwap|Microsoft.Online.SharePoint.PowerShell|
New-SPOMigrationEncryptionParameters|Microsoft.Online.SharePoint.PowerShell|
New-SPOMigrationPackage|Microsoft.Online.SharePoint.PowerShell|
New-SPOPublicCdnOrigin|Microsoft.Online.SharePoint.PowerShell|
New-SPOSdnProvider|Microsoft.Online.SharePoint.PowerShell|
New-SPOSite|Microsoft.Online.SharePoint.PowerShell|[spo site classic add](../cmd/spo/site/site-classic-add.md)
New-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
New-SPOSiteSharingReportJob|Microsoft.Online.SharePoint.PowerShell|
New-SPOTenantOrgRelation|Microsoft.Online.SharePoint.PowerShell|
Register-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Register-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite register](../cmd/spo/hubsite/hubsite-register.md)
Remove-SPODeletedSite|Microsoft.Online.SharePoint.PowerShell|[spo tenant recyclebinitem remove](../cmd/spo/tenant/tenant-recyclebinitem-remove.md)
Remove-SPOExternalUser|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOGeoAdministrator|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOHomeSite|Microsoft.Online.SharePoint.PowerShell|[spo homesite remove](../cmd/spo/homesite/homesite-remove.md)
Remove-SPOHubSiteAssociation|Microsoft.Online.SharePoint.PowerShell|[spo hubsite disconnect](../cmd/spo/hubsite/hubsite-disconnect.md)
Remove-SPOHubToHubAssociation|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOKnowledgeHubSite|Microsoft.Online.SharePoint.PowerShell|[spo knowledgehub remove](../cmd/spo/knowledgehub/knowledgehub-remove.md)
Remove-SPOMigrationJob|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOMultiGeoCompanyAllowedDataLocation|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOOrgAssetsLibrary|Microsoft.Online.SharePoint.PowerShell|[spo orgassetslibrary remove](../cmd/spo/orgassetslibrary/orgassetslibrary-remove.md)
Remove-SPOOrgNewsSite|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOPublicCdnOrigin|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSdnProvider|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSite|Microsoft.Online.SharePoint.PowerShell|[spo site remove](../cmd/spo/site/site-remove.md)
Remove-SPOSiteCollectionAppCatalog|Microsoft.Online.SharePoint.PowerShell|[spo site appcatalog remove](../cmd/spo/site/site-appcatalog-remove.md)
Remove-SPOSiteCollectionAppCatalogById|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign remove](../cmd/spo/sitedesign/sitedesign-remove.md)
Remove-SPOSiteDesignTask|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign task remove](../cmd/spo/sitedesign/sitedesign-task-remove.md)
Remove-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|[spo sitescript remove](../cmd/spo/sitescript/sitescript-remove.md)
Remove-SPOSiteSharingReportJob|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOSiteUserInvitations|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOStorageEntity|Microsoft.Online.SharePoint.PowerShell|[spo storageentity remove](../cmd/spo/storageentity/storageentity-remove.md)
Remove-SPOTenantCdnOrigin|Microsoft.Online.SharePoint.PowerShell|[spo cdn origin remove](../cmd/spo/cdn/cdn-origin-remove.md)
Remove-SPOTenantOrgRelation|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOTenantSyncClientRestriction|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOTheme|Microsoft.Online.SharePoint.PowerShell|[spo theme remove](../cmd/spo/theme/theme-remove.md)
Remove-SPOUser|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOUserInfo|Microsoft.Online.SharePoint.PowerShell|
Remove-SPOUserProfile|Microsoft.Online.SharePoint.PowerShell|
Repair-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Request-SPOPersonalSite|Microsoft.Online.SharePoint.PowerShell|
Request-SPOUpgradeEvaluationSite|Microsoft.Online.SharePoint.PowerShell|
Restore-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Restore-SPODeletedSite|Microsoft.Online.SharePoint.PowerShell|[spo tenant recyclebinitem restore](../cmd/spo/tenant/tenant-recyclebinitem-restore.md)
Revoke-SPOHubSiteRights|Microsoft.Online.SharePoint.PowerShell|[spo hubsite rights revoke](../cmd/spo/hubsite/hubsite-rights-revoke.md)
Revoke-SPOSiteDesignRights|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign rights revoke](../cmd/spo/sitedesign/sitedesign-rights-revoke.md)
Revoke-SPOTenantServicePrincipalPermission|Microsoft.Online.SharePoint.PowerShell|[spo serviceprincipal grant revoke](../cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md)
Revoke-SPOUserSession|Microsoft.Online.SharePoint.PowerShell|
Set-SPOBrowserIdleSignOut|Microsoft.Online.SharePoint.PowerShell|
Set-SPOBuiltDesignPackageVisibility|Microsoft.Online.SharePoint.PowerShell|
Set-SPOBuiltInDesignPackageVisibility|Microsoft.Online.SharePoint.PowerShell|
Set-SPOGeoStorageQuota|Microsoft.Online.SharePoint.PowerShell|
Set-SPOHideDefaultThemes|Microsoft.Online.SharePoint.PowerShell|[spo hidedefaultthemes set](../cmd/spo/hidedefaultthemes/hidedefaultthemes-set.md)
Set-SPOHomeSite|Microsoft.Online.SharePoint.PowerShell|[spo homesite set](../cmd/spo/homesite/homesite-set.md)
Set-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite set](../cmd/spo/hubsite/hubsite-set.md)
Set-SPOKnowledgeHubSite|Microsoft.Online.SharePoint.PowerShell|[spo knowledgehub set](../cmd/spo/knowledgehub/knowledgehub-set.md)
Set-SPOMigrationPackageAzureSource|Microsoft.Online.SharePoint.PowerShell|
Set-SPOMultiGeoCompanyAllowedDataLocation|Microsoft.Online.SharePoint.PowerShell|
Set-SPOMultiGeoExperience|Microsoft.Online.SharePoint.PowerShell|
Set-SPOOrgAssetsLibrary|Microsoft.Online.SharePoint.PowerShell|[spo orgassetslibrary add](../cmd/spo/orgassetslibrary/orgassetslibrary-add.md)
Set-SPOOrgNewsSite|Microsoft.Online.SharePoint.PowerShell|
Set-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Set-SPOSiteDesign|Microsoft.Online.SharePoint.PowerShell|[spo sitedesign set](../cmd/spo/sitedesign/sitedesign-set.md)
Set-SPOSiteGroup|Microsoft.Online.SharePoint.PowerShell|
Set-SPOSiteOffice365Group|Microsoft.Online.SharePoint.PowerShell|[spo site groupify](../cmd/spo/site/site-groupify.md)
Set-SPOSiteScript|Microsoft.Online.SharePoint.PowerShell|
Set-SPOSiteScriptPackage|Microsoft.Online.SharePoint.PowerShell|
Set-SPOStorageEntity|Microsoft.Online.SharePoint.PowerShell|[spo storageentity set](../cmd/spo/storageentity/storageentity-set.md)
Set-SPOStructuralNavigationCacheSiteState|Microsoft.Online.SharePoint.PowerShell|
Set-SPOStructuralNavigationCacheWebState|Microsoft.Online.SharePoint.PowerShell|
Set-SPOTenant|Microsoft.Online.SharePoint.PowerShell|[spo tenant settings set](../cmd/spo/tenant/tenant-settings-set.md)
Set-SPOTenantCdnEnabled|Microsoft.Online.SharePoint.PowerShell|[spo cdn set](../cmd/spo/cdn/cdn-set.md)
Set-SPOTenantCdnPolicy|Microsoft.Online.SharePoint.PowerShell|[spo cdn policy set](../cmd/spo/cdn/cdn-policy-set.md)
Set-SPOTenantContentTypeReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Set-SPOTenantSyncClientRestriction|Microsoft.Online.SharePoint.PowerShell|
Set-SPOTenantTaxonomyReplicationParameters|Microsoft.Online.SharePoint.PowerShell|
Set-SPOUnifiedGroup|Microsoft.Online.SharePoint.PowerShell|
Set-SPOUser|Microsoft.Online.SharePoint.PowerShell|
Set-SPOWebTheme|Microsoft.Online.SharePoint.PowerShell|
Start-SPOSiteContentMove|Microsoft.Online.SharePoint.PowerShell|
Start-SPOSiteRename|Microsoft.Online.SharePoint.PowerShell|[spo site rename](../cmd/spo/site/site-rename.md)
Start-SPOUnifiedGroupMove|Microsoft.Online.SharePoint.PowerShell|
Start-SPOUserAndContentMove|Microsoft.Online.SharePoint.PowerShell|
Stop-SPOSiteContentMove|Microsoft.Online.SharePoint.PowerShell|
Stop-SPOUserAndContentMove|Microsoft.Online.SharePoint.PowerShell|
Submit-SPOMigrationJob|Microsoft.Online.SharePoint.PowerShell|
Test-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Unregister-SPOHubSite|Microsoft.Online.SharePoint.PowerShell|[spo hubsite unregister](../cmd/spo/hubsite/hubsite-unregister.md)
Update-SPODataEncryptionPolicy|Microsoft.Online.SharePoint.PowerShell|
Update-UserType|Microsoft.Online.SharePoint.PowerShell|
Upgrade-SPOSite|Microsoft.Online.SharePoint.PowerShell|
Verify-SPOTenantOrgRelation|Microsoft.Online.SharePoint.PowerShell|
>> Start change here 
Add-PnPAlert|PnP.PowerShell|
Add-PnPApp|PnP.PowerShell|[spo app add](../cmd/spo/app/app-add.md)
Add-PnPApplicationCustomizer|PnP.PowerShell|[spo customaction add](../cmd/spo/customaction/customaction-add.md)
Add-PnPAzureADGroupMember|PnP.PowerShell|
Add-PnPAzureADGroupOwner
Add-PnPContentType
Add-PnPContentTypesFromContentTypeHub
Add-PnPContentTypeToDocumentSet
Add-PnPContentTypeToList
Add-PnPCustomAction
Add-PnPDataRowsToSiteTemplate
Add-PnPDocumentSet
Add-PnPEventReceiver
Add-PnPField
Add-PnPFieldFromXml
Add-PnPFieldToContentType
Add-PnPFile
Add-PnPFileToSiteTemplate
Add-PnPFolder
Add-PnPGroupMember
Add-PnPHtmlPublishingPageLayout
Add-PnPHubSiteAssociation
Add-PnPHubToHubAssociation
Add-PnPIndexedProperty
Add-PnPJavaScriptBlock
Add-PnPJavaScriptLink
Add-PnPListDesign
Add-PnPListFoldersToSiteTemplate
Add-PnPListItem
Add-PnPListItemAttachment
Add-PnPListItemComment
Add-PnPMasterPage
Add-PnPMicrosoft365GroupMember
Add-PnPMicrosoft365GroupOwner
Add-PnPMicrosoft365GroupToSite
Add-PnPNavigationNode
Add-PnPOrgAssetsLibrary
Add-PnPOrgNewsSite
Add-PnPPage
Add-PnPPageSection
Add-PnPPageTextPart
Add-PnPPageWebPart
Add-PnPPlannerBucket
Add-PnPPlannerRoster
Add-PnPPlannerRosterMember
Add-PnPPlannerTask
Add-PnPPublishingImageRendition
Add-PnPPublishingPage
Add-PnPPublishingPageLayout
Add-PnPRoleDefinition
Add-PnPSiteClassification
Add-PnPSiteCollectionAdmin
Add-PnPSiteCollectionAppCatalog
Add-PnPSiteDesign
Add-PnPSiteDesignFromWeb
Add-PnPSiteDesignTask
Add-PnPSiteScript
Add-PnPSiteScriptPackage
Add-PnPSiteTemplate
Add-PnPStoredCredential
Add-PnPTaxonomyField
Add-PnPTeamsChannel
Add-PnpTeamsChannelUser
Add-PnPTeamsTab
Add-PnPTeamsTeam
Add-PnPTeamsUser
Add-PnPTenantCdnOrigin
Add-PnPTenantSequence
Add-PnPTenantSequenceSite
Add-PnPTenantSequenceSubSite
Add-PnPTenantTheme
Add-PnPTermToTerm
Add-PnPView
Add-PnPViewsFromXML
Add-PnPVivaConnectionsDashboardACE
Add-PnPWebhookSubscription
Add-PnPWebPartToWebPartPage
Add-PnPWebPartToWikiPage
Add-PnPWikiPage
Approve-PnPTenantServicePrincipalPermissionRequest
Clear-PnPAzureADGroupMember
Clear-PnPAzureADGroupOwner
Clear-PnPDefaultColumnValues
Clear-PnPListItemAsRecord
Clear-PnPMicrosoft365GroupMember
Clear-PnPMicrosoft365GroupOwner
Clear-PnpRecycleBinItem
Clear-PnPTenantAppCatalogUrl
Clear-PnPTenantRecycleBinItem
Connect-PnPOnline
Convert-PnPFolderToSiteTemplate
Convert-PnPSiteTemplate
Convert-PnPSiteTemplateToMarkdown
ConvertTo-PnPClientSidePage
ConvertTo-PnPPage
Copy-PnPFile
Copy-PnPFolder
Copy-PnPList
Deny-PnPTenantServicePrincipalPermissionRequest
Disable-PnPFeature
Disable-PnPFlow
Disable-PnPPageScheduling
Disable-PnPSharingForNonOwnersOfSite
Disable-PnPSiteClassification
Disable-PnPTenantServicePrincipal
Disconnect-PnPOnline
Enable-PnPCommSite
Enable-PnPFeature
Enable-PnPFlow
Enable-PnPPageScheduling
Enable-PnPSiteClassification
Enable-PnPTenantServicePrincipal
Export-PnPFlow
Export-PnPListToSiteTemplate
Export-PnPPage
Export-PnPPageMapping
Export-PnPTaxonomy
Export-PnPTermGroupToXml
Export-PnPUserInfo
Export-PnPUserProfile
Find-PnPFile
Get-PnPAccessToken
Get-PnPAlert
Get-PnPApp
Get-PnPAppAuthAccessToken
Get-PnPAppErrors
Get-PnPAppInfo
Get-PnPApplicationCustomizer
Get-PnPAuditing
Get-PnPAuthenticationRealm
Get-PnPAvailableLanguage
Get-PnPAzureADApp
Get-PnPAzureADAppPermission
Get-PnPAzureADAppSitePermission
Get-PnPAzureADGroup
Get-PnPAzureADGroupMember
Get-PnPAzureADGroupOwner
Get-PnPAzureADUser
Get-PnPAzureCertificate
Get-PnPBrowserIdleSignout
Get-PnPBuiltInDesignPackageVisibility
Get-PnPBuiltInSiteTemplateSettings
Get-PnPChangeLog
Get-PnPCompatibleHubContentTypes
Get-PnPConnection
Get-PnPContentType
Get-PnPContentTypePublishingHubUrl
Get-PnPContentTypePublishingStatus
Get-PnPContext
Get-PnPCustomAction
Get-PnPDefaultColumnValues
Get-PnPDeletedMicrosoft365Group
Get-PnPDiagnostics
Get-PnPDisableSpacesActivation
Get-PnPDocumentSetTemplate
Get-PnPEventReceiver
Get-PnPException
Get-PnPExternalUser
Get-PnPFeature
Get-PnPField
Get-PnPFile
Get-PnPFileVersion
Get-PnPFlow
Get-PnPFlowRun
Get-PnPFolder
Get-PnPFolderItem
Get-PnPFooter
Get-PnPGraphAccessToken
Get-PnPGraphSubscription
Get-PnPGroup
Get-PnPGroupMember
Get-PnPGroupPermissions
Get-PnPHideDefaultThemes
Get-PnPHomePage
Get-PnPHomeSite
Get-PnPHubSite
Get-PnPHubSiteChild
Get-PnPIndexedPropertyKeys
Get-PnPInPlaceRecordsManagement
Get-PnPIsSiteAliasAvailable
Get-PnPJavaScriptLink
Get-PnPKnowledgeHubSite
Get-PnPLabel
Get-PnPList
Get-PnPListDesign
Get-PnPListInformationRightsManagement
Get-PnPListItem
Get-PnPListItemAttachments
Get-PnPListItemComment
Get-PnPListItemPermission
Get-PnPListPermissions
Get-PnPListRecordDeclaration
Get-PnPMasterPage
Get-PnPMessageCenterAnnouncement
Get-PnPMicrosoft365Group
Get-PnPMicrosoft365GroupMember
Get-PnPMicrosoft365GroupOwner
Get-PnPMicrosoft365GroupSettings
Get-PnPMicrosoft365GroupSettingTemplates
Get-PnPNavigationNode
Get-PnPOrgAssetsLibrary
Get-PnPOrgNewsSite
Get-PnPPage
Get-PnPPageComponent
Get-PnPPlannerBucket
Get-PnPPlannerConfiguration
Get-PnPPlannerPlan
Get-PnPPlannerRosterMember
Get-PnPPlannerRosterPlan
Get-PnPPlannerTask
Get-PnPPlannerUserPolicy
Get-PnPPowerPlatformEnvironment
Get-PnPPowerShellTelemetryEnabled
Get-PnPProperty
Get-PnPPropertyBag
Get-PnPPublishingImageRendition
Get-PnPRecycleBinItem
Get-PnPRequestAccessEmails
Get-PnPRoleDefinition
Get-PnPSearchConfiguration
Get-PnPSearchCrawlLog
Get-PnPSearchSettings
Get-PnPSensitivityLabel
Get-PnPServiceCurrentHealth
Get-PnPServiceHealthIssue
Get-PnPSharingForNonOwnersOfSite
Get-PnPSite
Get-PnPSiteClassification
Get-PnPSiteClosure
Get-PnPSiteCollectionAdmin
Get-PnPSiteCollectionAppCatalogs
Get-PnPSiteCollectionTermStore
Get-PnPSiteDesign
Get-PnPSiteDesignRights
Get-PnPSiteDesignRun
Get-PnPSiteDesignRunStatus
Get-PnPSiteDesignTask
Get-PnPSiteGroup
Get-PnPSitePolicy
Get-PnPSiteScript
Get-PnPSiteScriptFromList
Get-PnPSiteScriptFromWeb
Get-PnPSiteSearchQueryResults
Get-PnPSiteTemplate
Get-PnPSiteUserInvitations
Get-PnPStorageEntity
Get-PnPStoredCredential
Get-PnPStructuralNavigationCacheSiteState
Get-PnPStructuralNavigationCacheWebState
Get-PnPSubscribeSharePointNewsDigest
Get-PnPSubWeb
Get-PnPSyntexModel
Get-PnPSyntexModelPublication
Get-PnPTaxonomyItem
Get-PnPTaxonomySession
Get-PnPTeamsApp
Get-PnPTeamsChannel
Get-PnPTeamsChannelFilesFolder
Get-PnPTeamsChannelMessage
Get-PnPTeamsChannelMessageReply
Get-PnPTeamsChannelUser
Get-PnPTeamsPrimaryChannel
Get-PnPTeamsTab
Get-PnPTeamsTeam
Get-PnPTeamsUser
Get-PnPTemporarilyDisableAppBar
Get-PnPTenant
Get-PnPTenantAppCatalogUrl
Get-PnPTenantCdnEnabled
Get-PnPTenantCdnOrigin
Get-PnPTenantCdnPolicies
Get-PnPTenantDeletedSite
Get-PnPTenantId
Get-PnPTenantInstance
Get-PnPTenantRecycleBinItem
Get-PnPTenantSequence
Get-PnPTenantSequenceSite
Get-PnPTenantServicePrincipal
Get-PnPTenantServicePrincipalPermissionGrants
Get-PnPTenantServicePrincipalPermissionRequests
Get-PnPTenantSite
Get-PnPTenantSyncClientRestriction
Get-PnPTenantTemplate
Get-PnPTenantTheme
Get-PnPTerm
Get-PnPTermGroup
Get-PnPTermLabel
Get-PnPTermSet
Get-PnPTheme
Get-PnPTimeZoneId
Get-PnPUnifiedAuditLog
Get-PnPUPABulkImportStatus
Get-PnPUser
Get-PnPUserOneDriveQuota
Get-PnPUserProfileProperty
Get-PnPView
Get-PnPVivaConnectionsDashboardACE
Get-PnPWeb
Get-PnPWebHeader
Get-PnPWebhookSubscriptions
Get-PnPWebPart
Get-PnPWebPartProperty
Get-PnPWebPartXml
Get-PnPWebTemplates
Get-PnPWikiPageContent
Grant-PnPAzureADAppSitePermission
Grant-PnPHubSiteRights
Grant-PnPSiteDesignRights
Grant-PnPTenantServicePrincipalPermission
Import-PnPTaxonomy
Import-PnPTermGroupFromXml
Import-PnPTermSet
Install-PnPApp
Invoke-PnPBatch
Invoke-PnPGraphMethod
Invoke-PnPListDesign
Invoke-PnPQuery
Invoke-PnPSiteDesign
Invoke-PnPSiteScript
Invoke-PnPSiteSwap
Invoke-PnPSiteTemplate
Invoke-PnPSPRestMethod
Invoke-PnPTenantTemplate
Invoke-PnPWebAction
Measure-PnPList
Measure-PnPWeb
Move-PnPFile
Move-PnPFolder
Move-PnPListItemToRecycleBin
Move-PnPPageComponent
Move-PnpRecycleBinItem
New-PnPAzureADGroup
New-PnPAzureADUserTemporaryAccessPass
New-PnPAzureCertificate
New-PnPBatch
New-PnPExtensibilityHandlerObject
New-PnPGraphSubscription
New-PnPGroup
New-PnPList
New-PnPMicrosoft365Group
New-PnPMicrosoft365GroupSettings
New-PnPPersonalSite
New-PnPPlannerPlan
New-PnPSdnProvider
New-PnPSite
New-PnPSiteCollectionTermStore
New-PnPSiteGroup
New-PnPSiteTemplate
New-PnPSiteTemplateFromFolder
New-PnPTeamsApp
New-PnPTeamsTeam
New-PnPTenantSequence
New-PnPTenantSequenceCommunicationSite
New-PnPTenantSequenceTeamNoGroupSite
New-PnPTenantSequenceTeamNoGroupSubSite
New-PnPTenantSequenceTeamSite
New-PnPTenantSite
New-PnPTenantTemplate
New-PnPTerm
New-PnPTermGroup
New-PnPTermLabel
New-PnPTermSet
New-PnPUPABulkImportJob
New-PnPUser
New-PnPWeb
Publish-PnPApp
Publish-PnPCompanyApp
Publish-PnPContentType
Publish-PnPSyntexModel
Read-PnPSiteTemplate
Read-PnPTenantTemplate
Receive-PnPCopyMoveJobStatus
Register-PnPAppCatalogSite
Register-PnPAzureADApp
Register-PnPHubSite
Register-PnPManagementShellAccess
Remove-PnPAdaptiveScopeProperty
Remove-PnPAlert
Remove-PnPApp
Remove-PnPApplicationCustomizer
Remove-PnPAzureADApp
Remove-PnPAzureADAppSitePermission
Remove-PnPAzureADGroup
Remove-PnPAzureADGroupMember
Remove-PnPAzureADGroupOwner
Remove-PnPContentType
Remove-PnPContentTypeFromDocumentSet
Remove-PnPContentTypeFromList
Remove-PnPCustomAction
Remove-PnPDeletedMicrosoft365Group
Remove-PnPEventReceiver
Remove-PnPExternalUser
Remove-PnPField
Remove-PnPFieldFromContentType
Remove-PnPFile
Remove-PnPFileFromSiteTemplate
Remove-PnPFileVersion
Remove-PnPFlow
Remove-PnPFolder
Remove-PnPGraphSubscription
Remove-PnPGroup
Remove-PnPGroupMember
Remove-PnPHomeSite
Remove-PnPHubSiteAssociation
Remove-PnPHubToHubAssociation
Remove-PnPIndexedProperty
Remove-PnPJavaScriptLink
Remove-PnPKnowledgeHubSite
Remove-PnPList
Remove-PnPListDesign
Remove-PnPListItem
Remove-PnPListItemAttachment
Remove-PnPListItemComments
Remove-PnPMicrosoft365Group
Remove-PnPMicrosoft365GroupMember
Remove-PnPMicrosoft365GroupOwner
Remove-PnPMicrosoft365GroupSettings
Remove-PnPNavigationNode
Remove-PnPOrgAssetsLibrary
Remove-PnPOrgNewsSite
Remove-PnPPage
Remove-PnPPageComponent
Remove-PnPPlannerBucket
Remove-PnPPlannerPlan
Remove-PnPPlannerRoster
Remove-PnPPlannerRosterMember
Remove-PnPPlannerTask
Remove-PnPPropertyBagValue
Remove-PnPPublishingImageRendition
Remove-PnPRoleDefinition
Remove-PnPSdnProvider
Remove-PnPSearchConfiguration
Remove-PnPSiteClassification
Remove-PnPSiteCollectionAdmin
Remove-PnPSiteCollectionAppCatalog
Remove-PnPSiteCollectionTermStore
Remove-PnPSiteDesign
Remove-PnPSiteDesignTask
Remove-PnPSiteGroup
Remove-PnPSiteScript
Remove-PnPSiteUserInvitations
Remove-PnPStorageEntity
Remove-PnPStoredCredential
Remove-PnPTaxonomyItem
Remove-PnPTeamsApp
Remove-PnPTeamsChannel
Remove-PnPTeamsChannelUser
Remove-PnPTeamsTab
Remove-PnPTeamsTeam
Remove-PnPTeamsUser
Remove-PnPTenantCdnOrigin
Remove-PnPTenantDeletedSite
Remove-PnPTenantSite
Remove-PnPTenantSyncClientRestriction
Remove-PnPTenantTheme
Remove-PnPTerm
Remove-PnPTermGroup
Remove-PnPTermLabel
Remove-PnPUser
Remove-PnPUserInfo
Remove-PnPUserProfile
Remove-PnPView
Remove-PnPVivaConnectionsDashboardACE
Remove-PnPWeb
Remove-PnPWebhookSubscription
Remove-PnPWebPart
Remove-PnPWikiPage
Rename-PnPFile
Rename-PnPFolder
Rename-PnPTenantSite
Repair-PnPSite
Request-PnPAccessToken
Request-PnPPersonalSite
Request-PnPReIndexList
Request-PnPReIndexWeb
Request-PnPSyntexClassifyAndExtract
Reset-PnPFileVersion
Reset-PnPLabel
Reset-PnPMicrosoft365GroupExpiration
Reset-PnPUserOneDriveQuotaToDefault
Resolve-PnPFolder
Restart-PnPFlowRun
Restore-PnPDeletedMicrosoft365Group
Restore-PnPFileVersion
Restore-PnPRecycleBinItem
Restore-PnPTenantDeletedSite
Restore-PnPTenantRecycleBinItem
Revoke-PnPHubSiteRights
Revoke-PnPSiteDesignRights
Revoke-PnPTenantServicePrincipalPermission
Revoke-PnPUserSession
Save-PnPPageConversionLog
Save-PnPSiteTemplate
Save-PnPTenantTemplate
Send-PnPMail
Set-PnPAdaptiveScopeProperty
Set-PnPApplicationCustomizer
Set-PnPAppSideLoading
Set-PnPAuditing
Set-PnPAvailablePageLayouts
Set-PnPAzureADAppSitePermission
Set-PnPAzureADGroup
Set-PnPBrowserIdleSignout
Set-PnPBuiltInDesignPackageVisibility
Set-PnPBuiltInSiteTemplateSettings
Set-PnPContentType
Set-PnPContext
Set-PnPDefaultColumnValues
Set-PnPDefaultContentTypeToList
Set-PnPDefaultPageLayout
Set-PnPDisableSpacesActivation
Set-PnPDocumentSetField
Set-PnPField
Set-PnPFileCheckedIn
Set-PnPFileCheckedOut
Set-PnPFolderPermission
Set-PnPFooter
Set-PnPGraphSubscription
Set-PnPGroup
Set-PnPGroupPermissions
Set-PnPHideDefaultThemes
Set-PnPHomePage
Set-PnPHomeSite
Set-PnPHubSite
Set-PnPIndexedProperties
Set-PnPInPlaceRecordsManagement
Set-PnPKnowledgeHubSite
Set-PnPLabel
Set-PnPList
Set-PnPListInformationRightsManagement
Set-PnPListItem
Set-PnPListItemAsRecord
Set-PnPListItemPermission
Set-PnPListPermission
Set-PnPListRecordDeclaration
Set-PnPMasterPage
Set-PnPMessageCenterAnnouncementAsArchived
Set-PnPMessageCenterAnnouncementAsFavorite
Set-PnPMessageCenterAnnouncementAsNotArchived
Set-PnPMessageCenterAnnouncementAsNotFavorite
Set-PnPMessageCenterAnnouncementAsRead
Set-PnPMessageCenterAnnouncementAsUnread
Set-PnPMicrosoft365Group
Set-PnPMicrosoft365GroupSettings
Set-PnPMinimalDownloadStrategy
Set-PnPPage
Set-PnPPageTextPart
Set-PnPPageWebPart
Set-PnPPlannerBucket
Set-PnPPlannerConfiguration
Set-PnPPlannerPlan
Set-PnPPlannerTask
Set-PnPPlannerUserPolicy
Set-PnPPropertyBagValue
Set-PnPRequestAccessEmails
Set-PnPRoleDefinition
Set-PnPSearchConfiguration
Set-PnPSearchSettings
Set-PnPSite
Set-PnPSiteClosure
Set-PnPSiteDesign
Set-PnPSiteGroup
Set-PnPSitePolicy
Set-PnPSiteScript
Set-PnPSiteScriptPackage
Set-PnPSiteTemplateMetadata
Set-PnPStorageEntity
Set-PnPStructuralNavigationCacheSiteState
Set-PnPStructuralNavigationCacheWebState
Set-PnPSubscribeSharePointNewsDigest
Set-PnPTaxonomyFieldValue
Set-PnPTeamifyPromptHidden
Set-PnPTeamsChannel
Set-PnpTeamsChannelUser
Set-PnPTeamsTab
Set-PnPTeamsTeam
Set-PnPTeamsTeamArchivedState
Set-PnPTeamsTeamPicture
Set-PnPTemporarilyDisableAppBar
Set-PnPTenant
Set-PnPTenantAppCatalogUrl
Set-PnPTenantCdnEnabled
Set-PnPTenantCdnPolicy
Set-PnPTenantSite
Set-PnPTenantSyncClientRestriction
Set-PnPTerm
Set-PnPTermGroup
Set-PnPTermSet
Set-PnPTheme
Set-PnPTraceLog
Set-PnPUserOneDriveQuota
Set-PnPUserProfileProperty
Set-PnPView
Set-PnPWeb
Set-PnPWebHeader
Set-PnPWebhookSubscription
Set-PnPWebPartProperty
Set-PnPWebPermission
Set-PnPWebTheme
Set-PnPWikiPageContent
Stop-PnPFlowRun
Submit-PnPSearchQuery
Submit-PnPTeamsChannelMessage
Sync-PnPAppToTeams
Sync-PnPSharePointUserProfilesFromAzureActiveDirectory
Test-PnPListItemIsRecord
Test-PnPMicrosoft365GroupAliasIsUsed
Test-PnPSite
Test-PnPTenantTemplate
Uninstall-PnPApp
Unpublish-PnPApp
Unpublish-PnPContentType
Unpublish-PnPSyntexModel
Unregister-PnPHubSite
Update-PnPApp
Update-PnPSiteClassification
Update-PnPSiteDesignFromWeb
Update-PnPTeamsApp
Update-PnPTeamsUser
Update-PnPUserType
Update-PnPVivaConnectionsDashboardACE
>> end here
Approve-FlowApprovalRequest|Microsoft.PowerApps.PowerShell|
Deny-FlowApprovalRequest|Microsoft.PowerApps.PowerShell|
Disable-Flow|Microsoft.PowerApps.PowerShell|[flow disable](../cmd/flow/flow-disable.md)
Enable-Flow|Microsoft.PowerApps.PowerShell|[flow enable](../cmd/flow/flow-enable.md)
Get-Flow|Microsoft.PowerApps.PowerShell|[flow list](../cmd/flow/flow-list.md), [flow get](../cmd/flow/flow-get.md)
Get-FlowApproval|Microsoft.PowerApps.PowerShell|
Get-FlowApprovalRequest|Microsoft.PowerApps.PowerShell|
Get-FlowEnvironment|Microsoft.PowerApps.PowerShell|[flow environment list](../cmd/flow/environment/environment-list.md), [flow environment get](../cmd/flow/environment/environment-get.md)
Get-FlowOwnerRole|Microsoft.PowerApps.PowerShell|
Get-FlowRun|Microsoft.PowerApps.PowerShell|[flow run list](../cmd/flow/run/run-list.md), [flow run get](../cmd/flow/run/run-get.md)
Get-PowerApp|Microsoft.PowerApps.PowerShell|[pa app list](../cmd/pa/app/app-list.md), [pa app get](../cmd/pa/app/app-get.md)
Get-PowerAppConnection|Microsoft.PowerApps.PowerShell|
Get-PowerAppConnectionRoleAssignment|Microsoft.PowerApps.PowerShell|
Get-PowerAppConnector|Microsoft.PowerApps.PowerShell|
Get-PowerAppConnectorRoleAssignment|Microsoft.PowerApps.PowerShell|
Get-PowerAppEnvironment|Microsoft.PowerApps.PowerShell|[pa environment list](../cmd/pa/environment/environment-list.md), [pa environment get](../cmd/pa/environment/environment-get.md)
Get-PowerAppRoleAssignment|Microsoft.PowerApps.PowerShell|
Get-PowerAppsNotification|Microsoft.PowerApps.PowerShell|
Get-PowerAppVersion|Microsoft.PowerApps.PowerShell|
Publish-PowerApp|Microsoft.PowerApps.PowerShell|
Remove-Flow|Microsoft.PowerApps.PowerShell|[flow remove](../cmd/flow/flow-remove.md)
Remove-FlowOwnerRole|Microsoft.PowerApps.PowerShell|
Remove-PowerApp|Microsoft.PowerApps.PowerShell|[pa app remove](../cmd/pa/app/app-remove.md)
Remove-PowerAppConnection|Microsoft.PowerApps.PowerShell|
Remove-PowerAppConnectionRoleAssignment|Microsoft.PowerApps.PowerShell|
Remove-PowerAppConnector|Microsoft.PowerApps.PowerShell|
Remove-PowerAppConnectorRoleAssignment|Microsoft.PowerApps.PowerShell|
Remove-PowerAppRoleAssignment|Microsoft.PowerApps.PowerShell|
Restore-PowerAppVersion|Microsoft.PowerApps.PowerShell|
Set-FlowOwnerRole|Microsoft.PowerApps.PowerShell|
Set-PowerAppConnectionRoleAssignment|Microsoft.PowerApps.PowerShell|
Set-PowerAppConnectorRoleAssignment|Microsoft.PowerApps.PowerShell|
Set-PowerAppDisplayName|Microsoft.PowerApps.PowerShell|
Set-PowerAppRoleAssignment|Microsoft.PowerApps.PowerShell|
Add-ConnectorToBusinessDataGroup|Microsoft.PowerApps.Administration.PowerShell|
Add-CustomConnectorToPolicy|Microsoft.PowerApps.Administration.PowerShell|
Add-PowerAppsAccount|Microsoft.PowerApps.Administration.PowerShell|
Clear-AdminPowerAppApisToBypassConsent|Microsoft.PowerApps.Administration.PowerShell|
Clear-AdminPowerAppAsFeatured|Microsoft.PowerApps.Administration.PowerShell|
Clear-AdminPowerAppAsHero|Microsoft.PowerApps.Administration.PowerShell|
Disable-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|[flow disable](../cmd/flow/flow-disable.md)
Enable-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|[flow enable](../cmd/flow/flow-enable.md)
Get-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminFlowOwnerRole|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminFlowUserDetails|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerApp|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppCdsDatabaseCurrencies|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppCdsDatabaseLanguages|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnection|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnectionReferences|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnectionRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnector|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppConnectorRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppEnvironmentLocations|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppEnvironmentRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Get-AdminPowerAppsUserDetails|Microsoft.PowerApps.Administration.PowerShell|
Get-JwtToken|Microsoft.PowerApps.Administration.PowerShell|
Get-TenantDetailsFromGraph|Microsoft.PowerApps.Administration.PowerShell|
Get-UsersOrGroupsFromGraph|Microsoft.PowerApps.Administration.PowerShell|
InvokeApi|Microsoft.PowerApps.Administration.PowerShell|
New-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
New-AdminPowerAppCdsDatabase|Microsoft.PowerApps.Administration.PowerShell|
New-AdminPowerAppEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlow|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlowApprovals|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlowOwnerRole|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminFlowUserDetails|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerApp|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnection|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnectionRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnector|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppConnectorRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppEnvironmentRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-AdminPowerAppRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Remove-ConnectorFromBusinessDataGroup|Microsoft.PowerApps.Administration.PowerShell|
Remove-CustomConnectorFromPolicy|Microsoft.PowerApps.Administration.PowerShell|
Remove-LegacyCDSDatabase|Microsoft.PowerApps.Administration.PowerShell|
Remove-PowerAppsAccount|Microsoft.PowerApps.Administration.PowerShell|
ReplaceMacro|Microsoft.PowerApps.Administration.PowerShell|
Select-CurrentEnvironment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminDlpPolicy|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminFlowOwnerRole|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppApisToBypassConsent|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppAsFeatured|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppAsHero|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppConnectionRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppConnectorRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppEnvironmentDisplayName|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppEnvironmentRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppOwner|Microsoft.PowerApps.Administration.PowerShell|
Set-AdminPowerAppRoleAssignment|Microsoft.PowerApps.Administration.PowerShell|
Test-PowerAppsAccount|Microsoft.PowerApps.Administration.PowerShell|
Add-TeamUser|MicrosoftTeams|[teams user add](../cmd/aad/o365group/o365group-user-add.md)
Connect-MicrosoftTeams|MicrosoftTeams|[login](../cmd/login.md)
Disconnect-MicrosoftTeams|MicrosoftTeams|[logout](../cmd/logout.md)
Get-Team|MicrosoftTeams|[teams team list](../cmd/teams/team/team-list.md)
Get-TeamChannel|MicrosoftTeams|[teams channel list](../cmd/teams/channel/channel-list.md), [teams channel get](../cmd/teams/channel/channel-get.md)
Get-TeamFunSettings|MicrosoftTeams|[teams funsettings list](../cmd/teams/funsettings/funsettings-list.md)
Get-TeamGuestSettings|MicrosoftTeams|[teams guestsettings list](../cmd/teams/guestsettings/guestsettings-list.md)
Get-TeamMemberSettings|MicrosoftTeams|[teams membersettings list](../cmd/teams/membersettings/membersettings-list.md)
Get-TeamMessagingSettings|MicrosoftTeams|[teams messagingsettings list](../cmd/teams/messagingsettings/messagingsettings-list.md)
Get-TeamUser|MicrosoftTeams|[teams user list](../cmd/aad/o365group/o365group-user-list.md)
New-Team|MicrosoftTeams|[teams team add](../cmd/teams/team/team-add.md)
New-TeamChannel|MicrosoftTeams|[teams channel add](../cmd/teams/channel/channel-add.md)
Remove-Team|MicrosoftTeams|[teams team remove](../cmd/teams/team/team-remove.md)
Remove-TeamChannel|MicrosoftTeams|[teams channel remove](../cmd/teams/channel/channel-remove.md)
Remove-TeamUser|MicrosoftTeams|[teams user remove](../cmd/aad/o365group/o365group-user-remove.md)
Set-Team|MicrosoftTeams|[teams team set](../cmd/teams/team/team-set.md)
Set-TeamChannel|MicrosoftTeams|[teams channel set](../cmd/teams/channel/channel-set.md)
Set-TeamFunSettings|MicrosoftTeams|[teams funsettings set](../cmd/teams/funsettings/funsettings-set.md)
Set-TeamGuestSettings|MicrosoftTeams|[teams guestsettings set](../cmd/teams/guestsettings/guestsettings-set.md)
Set-TeamMemberSettings|MicrosoftTeams|[teams membersettings set](../cmd/teams/membersettings/membersettings-set.md)
Set-TeamMessagingSettings|MicrosoftTeams|[teams messagingsettings set](../cmd/teams/messagingsettings/messagingsettings-set.md)
Set-TeamPicture|MicrosoftTeams|
