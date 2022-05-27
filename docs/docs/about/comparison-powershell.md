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
Add-PnPAlert|PnP.PowerShell|
Add-PnPApp|PnP.PowerShell|[spo app add](../cmd/spo/app/app-add.md)
Add-PnPApplicationCustomizer
Add-PnPClientSidePage|PnP.PowerShell|[spo page add](../cmd/spo/page/page-add.md)
Add-PnPClientSidePageSection|PnP.PowerShell|[spo page section add](../cmd/spo/page/page-section-add.md)
Add-PnPClientSideText|PnP.PowerShell|
Add-PnPClientSideWebPart|PnP.PowerShell|[spo page clientsidewebpart add](../cmd/spo/page/page-clientsidewebpart-add.md)
Add-PnPContentType|PnP.PowerShell|[spo contenttype add](../cmd/spo/contenttype/contenttype-add.md)
Add-PnPContentTypeToDocumentSet|PnP.PowerShell|
Add-PnPContentTypeToList|PnP.PowerShell|[spo list contenttype add](../cmd/spo/list/list-contenttype-add.md)
Add-PnPCustomAction|PnP.PowerShell|[spo customaction add](../cmd/spo/customaction/customaction-add.md)
Add-PnPDataRowsToProvisioningTemplate|PnP.PowerShell|
Add-PnPDocumentSet|PnP.PowerShell|
Add-PnPEventReceiver|PnP.PowerShell|
Add-PnPField|PnP.PowerShell|
Add-PnPFieldFromXml|PnP.PowerShell|[spo field add](../cmd/spo/field/field-add.md)
Add-PnPFieldToContentType|PnP.PowerShell|[spo contenttype field set](../cmd/spo/contenttype/contenttype-field-set.md)
Add-PnPFile|PnP.PowerShell|[spo file add](../cmd/spo/file/file-add.md)
Add-PnPFileToProvisioningTemplate|PnP.PowerShell|
Add-PnPFolder|PnP.PowerShell|[spo folder add](../cmd/spo/folder/folder-add.md)
Add-PnPHtmlPublishingPageLayout|PnP.PowerShell|
Add-PnPHubSiteAssociation|PnP.PowerShell|[spo hubsite connect](../cmd/spo/hubsite/hubsite-connect.md)
Add-PnPIndexedProperty|PnP.PowerShell|
Add-PnPJavaScriptBlock|PnP.PowerShell|
Add-PnPJavaScriptLink|PnP.PowerShell|
Add-PnPListFoldersToProvisioningTemplate|PnP.PowerShell|
Add-PnPListItem|PnP.PowerShell|[spo listitem add](../cmd/spo/listitem/listitem-add.md)
Add-PnPMasterPage|PnP.PowerShell|
Add-PnPMicrosoft365GroupMember|PnP.PowerShell|[aad o365group user add](../cmd/aad/o365group/o365group-user-add.md)
Add-PnPMicrosoft365GroupOwner|PnP.PowerShell|[aad o365group user add](../cmd/aad/o365group/o365group-user-add.md)
Add-PnPMicrosoft365GroupToSite|PnP.PowerShell|
Add-PnPNavigationNode|PnP.PowerShell|[spo navigation node add](../cmd/spo/navigation/navigation-node-add.md)
Add-PnPOffice365GroupToSite|PnP.PowerShell|
Add-PnPOrgAssetsLibrary|PnP.PowerShell|[spo orgassetslibrary add](../cmd/spo/orgassetslibrary/orgassetslibrary-add.md)
Add-PnPOrgNewsSite|PnP.PowerShell|[spo orgnewssite set](../cmd/spo/orgnewssite/orgnewssite-set.md)
Add-PnPPlannerBucket|PnP.PowerShell|[planner bucket add](../cmd/planner/bucket/bucket-add.md)
Add-PnPPlannerTask|PnP.PowerShell|[planner task add](../cmd/planner/task/task-add.md)
Add-PnPProvisioningSequence|PnP.PowerShell|
Add-PnPProvisioningSite|PnP.PowerShell|
Add-PnPProvisioningTemplate|PnP.PowerShell|
Add-PnPPublishingImageRendition|PnP.PowerShell|
Add-PnPPublishingPage|PnP.PowerShell|
Add-PnPPublishingPageLayout|PnP.PowerShell|
Add-PnPRoleDefinition|PnP.PowerShell|
Add-PnPSiteClassification|PnP.PowerShell|
Add-PnPSiteCollectionAdmin|PnP.PowerShell|
Add-PnPSiteCollectionAppCatalog|PnP.PowerShell|[spo site appcatalog add](../cmd/spo/site/site-appcatalog-add.md)
Add-PnPSiteDesign|PnP.PowerShell|[spo sitedesign add](../cmd/spo/sitedesign/sitedesign-add.md)
Add-PnPSiteDesignTask|PnP.PowerShell|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
Add-PnPSiteScript|PnP.PowerShell|[spo sitescript add](../cmd/spo/sitescript/sitescript-add.md)
Add-PnPStoredCredential|PnP.PowerShell|
Add-PnPTaxonomyField|PnP.PowerShell|
Add-PnPTeamsChannel|PnP.PowerShell|[teams channel add](../cmd/teams/channel/channel-add.md)
Add-PnPTeamsChannelUser|PnP.PowerShell|[teams channel member add](../cmd/teams/channel/channel-member-add.md)
Add-PnPTeamsTab|PnP.PowerShell|[teams tab add](../cmd/teams/tab/tab-add.md)
Add-PnPTeamsTeam|PnP.PowerShell|[teams team add](../cmd/teams/team/team-add.md)
Add-PnPTeamsUser|PnP.PowerShell|[teams user add](../cmd/aad/o365group/o365group-user-add.md)
Add-PnPTenantCdnOrigin|PnP.PowerShell|[spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md)
Add-PnPTenantSequence|PnP.PowerShell|
Add-PnPTenantSequenceSite|PnP.PowerShell|
Add-PnPTenantSequenceSubSite|PnP.PowerShell|
Add-PnPTenantTheme|PnP.PowerShell|[spo theme set](../cmd/spo/theme/theme-set.md)
Add-PnPUnifiedGroupMember|PnP.PowerShell|[aad o365group user add](../cmd/aad/o365group/o365group-user-add.md)
Add-PnPUnifiedGroupOwner|PnP.PowerShell|[aad o365group user add](../cmd/aad/o365group/o365group-user-add.md)
Add-PnPUserToGroup|PnP.PowerShell|
Add-PnPView|PnP.PowerShell|
Add-PnPView|PnP.PowerShell|[spo list view add](../cmd/spo/list/list-view-add.md)
Add-PnPWebhookSubscription|PnP.PowerShell|[spo list webhook add](../cmd/spo/list/list-webhook-add.md)
Add-PnPWebPartToWebPartPage|PnP.PowerShell|
Add-PnPWebPartToWikiPage|PnP.PowerShell|
Add-PnPWikiPage|PnP.PowerShell|
Add-PnPWorkflowDefinition|PnP.PowerShell|
Add-PnPWorkflowSubscription|PnP.PowerShell|
Apply-PnPProvisioningHierarchy|PnP.PowerShell|
Apply-PnPProvisioningTemplate|PnP.PowerShell|
Apply-PnPTenantTemplate|PnP.PowerShell|
Approve-PnPTenantServicePrincipalPermissionRequest|PnP.PowerShell|[spo serviceprincipal permissionrequest approve](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md)
Clear-PnPDefaultColumnValues|PnP.PowerShell|
Clear-PnPListItemAsRecord|PnP.PowerShell|[spo listitem record undeclare](../cmd/spo/listitem/listitem-record-undeclare.md)
Clear-PnPMicrosoft365GroupMember|PnP.PowerShell|
Clear-PnPMicrosoft365GroupMember|PnP.PowerShell|
Clear-PnPMicrosoft365GroupOwner|PnP.PowerShell|
Clear-PnPRecycleBinItem|PnP.PowerShell|
Clear-PnPTenantAppCatalogUrl|PnP.PowerShell|
Clear-PnPTenantRecycleBinItem|PnP.PowerShell|[spo tenant recyclebinitem remove](../cmd/spo/tenant/tenant-recyclebinitem-remove.md)
Clear-PnPUnifiedGroupOwner|PnP.PowerShell|
Connect-PnPHubSite|PnP.PowerShell|[spo hubsite connect](../cmd/spo/hubsite/hubsite-connect.md)
Connect-PnPMicrosoftGraph|PnP.PowerShell|[login](../cmd/login.md)
Connect-PnPOnline|PnP.PowerShell|[login](../cmd/login.md)
Convert-PnPFolderToProvisioningTemplate|PnP.PowerShell|
Convert-PnPProvisioningTemplate|PnP.PowerShell|
ConvertTo-PnPClientSidePage|PnP.PowerShell|
Copy-PnPFile|PnP.PowerShell|[spo file copy](../cmd/spo/file/file-copy.md), [spo folder copy](../cmd/spo/folder/folder-copy.md)
Copy-PnPItemProxy|PnP.PowerShell|
Deny-PnPTenantServicePrincipalPermissionRequest|PnP.PowerShell|[spo serviceprincipal permissionrequest deny](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md)
Disable-PnPFeature|PnP.PowerShell|[spo feature disable](../cmd/spo/feature/feature-disable.md)
Disable-PnPInPlaceRecordsManagementForSite|PnP.PowerShell|[spo site inplacerecordsmanagement set](../cmd/spo/site/site-inplacerecordsmanagement-set.md)
Disable-PnPPowerShellTelemetry|PnP.PowerShell|
Disable-PnPResponsiveUI|PnP.PowerShell|
Disable-PnPSharingForNonOwnersOfSite|PnP.PowerShell|
Disable-PnPSiteClassification|PnP.PowerShell|[aad siteclassification disable](../cmd/aad/siteclassification/siteclassification-disable.md)
Disable-PnPTenantServicePrincipal|PnP.PowerShell|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Disconnect-PnPHubSite|PnP.PowerShell|[spo hubsite disconnect](../cmd/spo/hubsite/hubsite-disconnect.md)
Disconnect-PnPOnline|PnP.PowerShell|[logout](../cmd/logout.md)
Enable-PnPCommSite|PnP.PowerShell|[spo site commsite enable](../cmd/spo/site/site-commsite-enable.md)
Enable-PnPFeature|PnP.PowerShell|[spo feature enable](../cmd/spo/feature/feature-enable.md)
Enable-PnPInPlaceRecordsManagementForSite|PnP.PowerShell|
Enable-PnPPowerShellTelemetry|PnP.PowerShell|
Enable-PnPResponsiveUI|PnP.PowerShell|
Enable-PnPSiteClassification|PnP.PowerShell|[aad siteclassification enable](../cmd/aad/siteclassification/siteclassification-enable.md)
Enable-PnPTenantServicePrincipal|PnP.PowerShell|[spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md)
Ensure-PnPFolder|PnP.PowerShell|
Execute-PnPQuery|PnP.PowerShell|
Export-PnPClientSidePage|PnP.PowerShell|
Export-PnPClientSidePageMapping|PnP.PowerShell|
Export-PnPListToProvisioningTemplate|PnP.PowerShell|
Export-PnPTaxonomy|PnP.PowerShell|
Export-PnPTermGroupToXml|PnP.PowerShell|
Find-PnPFile|PnP.PowerShell|
Get-PnPAADUser|PnP.PowerShell|[aad user get](../cmd/aad/user/user-get.md), [aad user list](../cmd/aad/user/user-list.md)
Get-PnPAccessToken|PnP.PowerShell|[util accesstoken get](../cmd/util/accesstoken/accesstoken-get.md)
Get-PnPAlert|PnP.PowerShell|
Get-PnPApp|PnP.PowerShell|[spo app get](../cmd/spo/app/app-get.md), [spo app list](../cmd/spo/app/app-list.md)
Get-PnPAppAuthAccessToken|PnP.PowerShell|
Get-PnPAppInstance|PnP.PowerShell|
Get-PnPApplicationCustomizer|PnP.PowerShell|
Get-PnPAuditing|PnP.PowerShell|
Get-PnPAuthenticationRealm|PnP.PowerShell|
Get-PnPAvailableClientSideComponents|PnP.PowerShell|
Get-PnPAvailableLanguage|PnP.PowerShell|
Get-PnPAzureADManifestKeyCredentials|PnP.PowerShell|
Get-PnPAzureCertificate|PnP.PowerShell|
Get-PnPClientSideComponent|PnP.PowerShell|
Get-PnPClientSidePage|PnP.PowerShell|[spo page get](../cmd/spo/page/page-get.md), [spo page control list](../cmd/spo/page/page-control-list.md), [spo page control get](../cmd/spo/page/page-control-get.md), [spo page section get](../cmd/spo/page/page-section-get.md), [spo page section list](../cmd/spo/page/page-section-list.md), [spo page column get](../cmd/spo/page/page-column-get.md), [spo page column list](../cmd/spo/page/page-column-list.md), [spo page text add](../cmd/spo/page/page-text-add.md)
Get-PnPConnection|PnP.PowerShell|
Get-PnPContentType|PnP.PowerShell|[spo contenttype get](../cmd/spo/contenttype/contenttype-get.md), [spo list contenttype list](../cmd/spo/list/list-contenttype-list.md)
Get-PnPContentTypePublishingHubUrl|PnP.PowerShell|[spo contenttypehub get](../cmd/spo/contenttypehub/contenttypehub-get.md)
Get-PnPContext|PnP.PowerShell|
Get-PnPCustomAction|PnP.PowerShell|[spo customaction get](../cmd/spo/customaction/customaction-get.md), [spo customaction list](../cmd/spo/customaction/customaction-list.md)
Get-PnPDefaultColumnValues|PnP.PowerShell|
Get-PnPDeletedMicrosoft365Group|PnP.PowerShell|[aad o365group list](../cmd/aad/o365group/o365group-list.md)
Get-PnPDeletedUnifiedGroup|PnP.PowerShell|[aad o365group list](../cmd/aad/o365group/o365group-list.md)
Get-PnPDocumentSetTemplate|PnP.PowerShell|
Get-PnPEventReceiver|PnP.PowerShell|
Get-PnPEventReceiver|PnP.PowerShell|[spo eventreceiver get](../cmd/spo/eventreceiver/eventreceiver-get.md)
Get-PnPException|PnP.PowerShell|
Get-PnPFeature|PnP.PowerShell|[spo feature list](../cmd/spo/feature/feature-list.md)
Get-PnPField|PnP.PowerShell|[spo field get](../cmd/spo/field/field-get.md)
Get-PnPFile|PnP.PowerShell|[spo file get](../cmd/spo/file/file-get.md), [spo file list](../cmd/spo/file/file-list.md)
Get-PnPFileVersion|PnP.PowerShell|
Get-PnPFlowRun|PnP.PowerShell|[flow run get](../cmd/flow/run/run-get.md), [flow run list](../cmd/flow/run/run-list.md)
Get-PnPFolder|PnP.PowerShell|[spo folder get](../cmd/spo/folder/folder-get.md), [spo folder list](../cmd/spo/folder/folder-list.md)
Get-PnPFolderItem|PnP.PowerShell|
Get-PnPFooter|PnP.PowerShell|
Get-PnPGraphAccessToken|PnP.PowerShell|[util accesstoken get](../cmd/util/accesstoken/accesstoken-get.md)
Get-PnPGraphSubscription|PnP.PowerShell|
Get-PnPGroup|PnP.PowerShell|[spo group get](../cmd/spo/group/group-get.md), [spo group list](../cmd/spo/group/group-list.md)
Get-PnPGroupMembers|PnP.PowerShell|
Get-PnPGroupPermissions|PnP.PowerShell|
Get-PnPHealthScore|PnP.PowerShell|
Get-PnPHideDefaultThemes|PnP.PowerShell|[spo hidedefaultthemes get](../cmd/spo/hidedefaultthemes/hidedefaultthemes-get.md)
Get-PnPHomePage|PnP.PowerShell|
Get-PnPHomeSite|PnP.PowerShell|[spo homesite get](../cmd/spo/homesite/homesite-get.md)
Get-PnPHubSite|PnP.PowerShell|[spo hubsite get](../cmd/spo/hubsite/hubsite-get.md), [spo hubsite list](../cmd/spo/hubsite/hubsite-list.md)
Get-PnPHubSiteChild|PnP.PowerShell|
Get-PnPIndexedPropertyKeys|PnP.PowerShell|
Get-PnPInPlaceRecordsManagement|PnP.PowerShell|
Get-PnPIsSiteAliasAvailable|PnP.PowerShell|
Get-PnPJavaScriptLink|PnP.PowerShell|
Get-PnPKnowledgeHubSite|PnP.PowerShell|
Get-PnPLabel|PnP.PowerShell|[spo list label get](../cmd/spo/list/list-label-get.md)
Get-PnPList|PnP.PowerShell|[spo list get](../cmd/spo/list/list-get.md), [spo list list](../cmd/spo/list/list-list.md)
Get-PnPListInformationRightsManagement|PnP.PowerShell|
Get-PnPListItem|PnP.PowerShell|[spo listitem get](../cmd/spo/listitem/listitem-get.md), [spo listitem list](../cmd/spo/listitem/listitem-list.md)
Get-PnPListRecordDeclaration|PnP.PowerShell|
Get-PnPManagementApiAccessToken|PnP.PowerShell|[util accesstoken get](../cmd/util/accesstoken/accesstoken-get.md)
Get-PnPMasterPage|PnP.PowerShell|
Get-PnPMicrosoft365Group|PnP.PowerShell|[aad o365group get](../cmd/aad/o365group/o365group-get.md)
Get-PnPMicrosoft365GroupMembers|PnP.PowerShell|[aad o365group user list](../cmd/aad/o365group/o365group-user-list.md)
Get-PnPMicrosoft365GroupOwners|PnP.PowerShell|[aad o365group user list](../cmd/aad/o365group/o365group-user-list.md)
Get-PnPNavigationNode|PnP.PowerShell|[spo navigation node list](../cmd/spo/navigation/navigation-node-list.md)
Get-PnPOffice365CurrentServiceStatus|PnP.PowerShell|
Get-PnPOffice365HistoricalServiceStatus|PnP.PowerShell|
Get-PnPOffice365ServiceMessage|PnP.PowerShell|
Get-PnPOffice365Services|PnP.PowerShell|
Get-PnPOfficeManagementApiAccessToken|PnP.PowerShell|[util accesstoken get](../cmd/util/accesstoken/accesstoken-get.md)
Get-PnPOrgAssetsLibrary|PnP.PowerShell|[spo orgassetslibrary list](../cmd/spo/orgassetslibrary/orgassetslibrary-list.md)
Get-PnPOrgNewsSite|PnP.PowerShell|[spo orgnewssite list](../cmd/spo/orgnewssite/orgnewssite-list.md)
Get-PnPPlannerBucket|PnP.PowerShell|[planner bucket get](../cmd/planner/bucket/bucket-get.md), [planner bucket list](../cmd/planner/bucket/bucket-list.md)
Get-PnPPlannerTask|PnP.PowerShell|[planner task get](../cmd/planner/task/task-get.md), [planner task list](../cmd/planner/task/task-list.md)
Get-PnPPowerShellTelemetryEnabled|PnP.PowerShell|
Get-PnPProperty|PnP.PowerShell|
Get-PnPPropertyBag|PnP.PowerShell|[spo propertybag get](../cmd/spo/propertybag/propertybag-get.md), [spo propertybag list](../cmd/spo/propertybag/propertybag-list.md)
Get-PnPProvisioningSequence|PnP.PowerShell|
Get-PnPProvisioningSite|PnP.PowerShell|
Get-PnPProvisioningTemplate|PnP.PowerShell|
Get-PnPPublishingImageRendition|PnP.PowerShell|
Get-PnPRecycleBinItem|PnP.PowerShell|
Get-PnPRequestAccessEmails|PnP.PowerShell|
Get-PnPRoleDefinition|PnP.PowerShell|
Get-PnPSearchConfiguration|PnP.PowerShell|
Get-PnPSearchCrawlLog|PnP.PowerShell|
Get-PnPSearchSettings|PnP.PowerShell|
Get-PnPSharingForNonOwnersOfSite
Get-PnPSite|PnP.PowerShell|[spo site get](../cmd/spo/site/site-get.md), [spo site list](../cmd/spo/site/site-list.md)
Get-PnPSiteClassification|PnP.PowerShell|[aad siteclassification get](../cmd/aad/siteclassification/siteclassification-get.md)
Get-PnPSiteClosure|PnP.PowerShell|
Get-PnPSiteCollectionAdmin|PnP.PowerShell|
Get-PnPSiteCollectionTermStore|PnP.PowerShell|
Get-PnPSiteDesign|PnP.PowerShell|[spo sitedesign get](../cmd/spo/sitedesign/sitedesign-get.md), [spo sitedesign list](../cmd/spo/sitedesign/sitedesign-list.md)
Get-PnPSiteDesignRights|PnP.PowerShell|[spo sitedesign rights list](../cmd/spo/sitedesign/sitedesign-rights-list.md)
Get-PnPSiteDesignRun|PnP.PowerShell|[spo sitedesign run list](../cmd/spo/sitedesign/sitedesign-run-list.md)
Get-PnPSiteDesignRunStatus|PnP.PowerShell|[spo sitedesign run status get](../cmd/spo/sitedesign/sitedesign-run-status-get.md)
Get-PnPSiteDesignTask|PnP.PowerShell|[spo sitedesign task get](../cmd/spo/sitedesign/sitedesign-task-get.md), [spo sitedesign task list](../cmd/spo/sitedesign/sitedesign-task-list.md)
Get-PnPSitePolicy|PnP.PowerShell|
Get-PnPSiteScript|PnP.PowerShell|[spo sitescript get](../cmd/spo/sitescript/sitescript-get.md), [spo sitescript list](../cmd/spo/sitescript/sitescript-list.md)
Get-PnPSiteScriptFromList|PnP.PowerShell|[spo list sitescript get](../cmd/spo/list/list-sitescript-get.md)
Get-PnPSiteScriptFromWeb|PnP.PowerShell|
Get-PnPSiteSearchQueryResults|PnP.PowerShell|
Get-PnPStorageEntity|PnP.PowerShell|[spo storageentity get](../cmd/spo/storageentity/storageentity-get.md), [spo storageentity list](../cmd/spo/storageentity/storageentity-list.md)
Get-PnPStoredCredential|PnP.PowerShell|
Get-PnPSubWebs|PnP.PowerShell|
Get-PnPTaxonomyItem|PnP.PowerShell|
Get-PnPTaxonomySession|PnP.PowerShell|
Get-PnPTeamsApp|PnP.PowerShell|[teams app list](../cmd/teams/app/app-list.md)
Get-PnPTeamsChannel|PnP.PowerShell|[teams channel get](../cmd/teams/channel/channel-get.md), [teams channel list](../cmd/teams/channel/channel-list.md)
Get-PnPTeamsChannelMessage|PnP.PowerShell|[teams message get](../cmd/teams/message/message-get.md), [teams message list](../cmd/teams/message/message-list.md)
Get-PnPTeamsChannelMessageReply|PnP.PowerShell|[teams message reply list](../cmd/teams/message/message-reply-list.md)
Get-PnPTeamsChannelUser|PnP.PowerShell|[teams channel member list](../cmd/teams/channel/channel-member-list.md)
Get-PnPTeamsTab|PnP.PowerShell|[teams tab list](../cmd/teams/tab/tab-list.md)
Get-PnPTeamsTeam|PnP.PowerShell|[teams team list](../cmd/teams/team/team-list.md)
Get-PnPTeamsUser|PnP.PowerShell|[teams user list](../cmd/teams/user/user-list.md)
Get-PnPTenant|PnP.PowerShell|[spo tenant settings list](../cmd/spo/tenant/tenant-settings-list.md)
Get-PnPTenantAppCatalogUrl|PnP.PowerShell|[spo tenant appcatalogurl get](../cmd/spo/tenant/tenant-appcatalogurl-get.md)
Get-PnPTenantCdnEnabled|PnP.PowerShell|[spo cdn get](../cmd/spo/cdn/cdn-get.md)
Get-PnPTenantCdnOrigin|PnP.PowerShell|[spo cdn origin list](../cmd/spo/cdn/cdn-origin-list.md)
Get-PnPTenantCdnPolicies|PnP.PowerShell|[spo cdn policy list](../cmd/spo/cdn/cdn-policy-list.md)
Get-PnPTenantId|PnP.PowerShell|[tenant id get](../cmd/tenant/id/id-get.md)
Get-PnPTenantRecycleBinItem|PnP.PowerShell|[spo tenant recyclebinitem list](../cmd/spo/tenant/tenant-recyclebinitem-list.md)
Get-PnPTenantSequence|PnP.PowerShell|
Get-PnPTenantSequenceSite|PnP.PowerShell|
Get-PnPTenantServicePrincipal|PnP.PowerShell|
Get-PnPTenantServicePrincipalPermissionGrants|PnP.PowerShell|[spo serviceprincipal grant list](../cmd/spo/serviceprincipal/serviceprincipal-grant-list.md)
Get-PnPTenantServicePrincipalPermissionRequests|PnP.PowerShell|[spo serviceprincipal permissionrequest list](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-list.md)
Get-PnPTenantSite|PnP.PowerShell|[spo site get](../cmd/spo/site/site-get.md), [spo site classic list](../cmd/spo/site/site-classic-list.md)
Get-PnPTenantSyncClientRestriction|PnP.PowerShell|
Get-PnPTenantTemplate|PnP.PowerShell|
Get-PnPTenantTheme|PnP.PowerShell|[spo theme get](../cmd/spo/theme/theme-get.md), [spo theme list](../cmd/spo/theme/theme-list.md)
Get-PnPTerm|PnP.PowerShell|[spo term get](../cmd/spo/term/term-get.md), [spo term list](../cmd/spo/term/term-list.md)
Get-PnPTermGroup|PnP.PowerShell|[spo term group get](../cmd/spo/term/term-group-get.md), [spo term group list](../cmd/spo/term/term-group-list.md)
Get-PnPTermSet|PnP.PowerShell|[spo term set get](../cmd/spo/term/term-set-get.md), [spo term set list](../cmd/spo/term/term-set-list.md)
Get-PnPTheme|PnP.PowerShell|
Get-PnPTimeZoneId|PnP.PowerShell|
Get-PnPUnifiedAuditLog|PnP.PowerShell|
Get-PnPUnifiedGroup|PnP.PowerShell|[aad o365group get](../cmd/aad/o365group/o365group-get.md), [aad o365group list](../cmd/aad/o365group/o365group-list.md)
Get-PnPUPABulkImportStatus|PnP.PowerShell|
Get-PnPUser|PnP.PowerShell|[spo user get](../cmd/spo/user/user-get.md), [spo user list](../cmd/spo/user/user-list.md)
Get-PnPUserOneDriveQuota|PnP.PowerShell|
Get-PnPUserProfileProperty|PnP.PowerShell|
Get-PnPView|PnP.PowerShell|[spo list view get](../cmd/spo/list/list-view-get.md), [spo list view list](../cmd/spo/list/list-view-list.md)
Get-PnPWeb|PnP.PowerShell|[spo web get](../cmd/spo/web/web-get.md), [spo web list](../cmd/spo/web/web-list.md)
Get-PnPWebhookSubscriptions|PnP.PowerShell|[spo list webhook get](../cmd/spo/list/list-webhook-get.md), [spo list webhook list](../cmd/spo/list/list-webhook-list.md)
Get-PnPWebPart|PnP.PowerShell|
Get-PnPWebPartProperty|PnP.PowerShell|
Get-PnPWebPartXml|PnP.PowerShell|
Get-PnPWebTemplates|PnP.PowerShell|
Get-PnPWikiPageContent|PnP.PowerShell|
Get-PnPWorkflowDefinition|PnP.PowerShell|
Get-PnPWorkflowInstance|PnP.PowerShell|
Get-PnPWorkflowSubscription|PnP.PowerShell|
Grant-PnPHubSiteRights|PnP.PowerShell|[spo hubsite rights grant](../cmd/spo/hubsite/hubsite-rights-grant.md)
Grant-PnPSiteDesignRights|PnP.PowerShell|[spo sitedesign rights grant](../cmd/spo/sitedesign/sitedesign-rights-grant.md)
Grant-PnPTenantServicePrincipalPermission|PnP.PowerShell|[aad oauth2grant add](../cmd/aad/oauth2grant/oauth2grant-add.md)
Import-PnPAppPackage|PnP.PowerShell|
Import-PnPTaxonomy|PnP.PowerShell|
Import-PnPTermGroupFromXml|PnP.PowerShell|
Import-PnPTermSet|PnP.PowerShell|
Initialize-PnPPowerShellAuthentication|PnP.PowerShell|
Install-PnPApp|PnP.PowerShell|[spo app install](../cmd/spo/app/app-install.md)
Install-PnPSolution|PnP.PowerShell|
Invoke-PnPQuery|PnP.PowerShell|
Invoke-PnPSearchQuery|PnP.PowerShell|[spo search](../cmd/spo/spo-search.md)
Invoke-PnPSiteDesign|PnP.PowerShell|[spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md)
Invoke-PnPSPRestMethod|PnP.PowerShell|
Invoke-PnPWebAction|PnP.PowerShell|
Load-PnPProvisioningTemplate|PnP.PowerShell|
Measure-PnPList|PnP.PowerShell|
Measure-PnPResponseTime|PnP.PowerShell|
Measure-PnPWeb|PnP.PowerShell|
Move-PnPClientSideComponent|PnP.PowerShell|
Move-PnPFile|PnP.PowerShell|[spo file move](../cmd/spo/file/file-copy.md)
Move-PnPFolder|PnP.PowerShell|[spo folder move](../cmd/spo/folder/folder-move.md)
Move-PnPItemProxy|PnP.PowerShell|
Move-PnPListItemToRecycleBin|PnP.PowerShell|
Move-PnPRecycleBinItem|PnP.PowerShell|
New-PnPAzureCertificate|PnP.PowerShell|
New-PnPExtensibilityHandlerObject|PnP.PowerShell|
New-PnPGraphSubscription|PnP.PowerShell|
New-PnPGroup|PnP.PowerShell|
New-PnPList|PnP.PowerShell|[spo list add](../cmd/spo/list/list-add.md)
New-PnPMicrosoft365Group|PnP.PowerShell|[aad o365group add](../cmd/aad/o365group/o365group-add.md)
New-PnPPersonalSite|PnP.PowerShell|
New-PnPPlannerPlan|PnP.PowerShell|[planner plan add](../cmd/planner/plan/plan-add.md)
New-PnPProvisioningCommunicationSite|PnP.PowerShell|
New-PnPProvisioningHierarchy|PnP.PowerShell|
New-PnPProvisioningSequence|PnP.PowerShell|
New-PnPProvisioningTeamNoGroupSite|PnP.PowerShell|
New-PnPProvisioningTeamNoGroupSubSite|PnP.PowerShell|
New-PnPProvisioningTeamSite|PnP.PowerShell|
New-PnPProvisioningTemplate|PnP.PowerShell|
New-PnPProvisioningTemplateFromFolder|PnP.PowerShell|
New-PnPSite|PnP.PowerShell|[spo site add](../cmd/spo/site/site-add.md)
New-PnPTeamsApp|PnP.PowerShell|
New-PnPTeamsTeam|PnP.PowerShell|[teams team add](../cmd/teams/team/team-add.md)
New-PnPTenantSequence|PnP.PowerShell|
New-PnPTenantSequenceCommunicationSite|PnP.PowerShell|
New-PnPTenantSequenceTeamNoGroupSite|PnP.PowerShell|
New-PnPTenantSequenceTeamNoGroupSubSite|PnP.PowerShell|
New-PnPTenantSequenceTeamSite|PnP.PowerShell|
New-PnPTenantSite|PnP.PowerShell|[spo site classic add](../cmd/spo/site/site-classic-add.md)
New-PnPTenantTemplate|PnP.PowerShell|
New-PnPTerm|PnP.PowerShell|[spo term add](../cmd/spo/term/term-add.md)
New-PnPTermGroup|PnP.PowerShell|[spo term group add](../cmd/spo/term/term-group-add.md)
New-PnPTermLabel|PnP.PowerShell|
New-PnPTermSet|PnP.PowerShell|[spo term set add](../cmd/spo/term/term-set-add.md)
New-PnPUnifiedGroup|PnP.PowerShell|[aad o365group add](../cmd/aad/o365group/o365group-add.md)
New-PnPUPABulkImportJob|PnP.PowerShell|
New-PnPUser|PnP.PowerShell|
New-PnPWeb|PnP.PowerShell|[spo web add](../cmd/spo/web/web-add.md)
Publish-PnPApp|PnP.PowerShell|[spo app deploy](../cmd/spo/app/app-deploy.md)
Read-PnPProvisioningHierarchy|PnP.PowerShell|
Read-PnPProvisioningTemplate|PnP.PowerShell|
Read-PnPTenantTemplate|PnP.PowerShell|
Register-PnPAppCatalogSite|PnP.PowerShell|[spo site appcatalog add](../cmd/spo/site/site-appcatalog-add.md)
Register-PnPHubSite|PnP.PowerShell|[spo hubsite register](../cmd/spo/hubsite/hubsite-register.md)
Remove-PnPAlert|PnP.PowerShell|
Remove-PnPApp|PnP.PowerShell|[spo app remove](../cmd/spo/app/app-remove.md)
Remove-PnPApplicationCustomizer|PnP.PowerShell|
Remove-PnPClientSideComponent|PnP.PowerShell|
Remove-PnPClientSidePage|PnP.PowerShell|[spo page remove](../cmd/spo/page/page-remove.md)
Remove-PnPContentType|PnP.PowerShell|[spo contenttype remove](../cmd/spo/contenttype/contenttype-remove.md)
Remove-PnPContentTypeFromDocumentSet|PnP.PowerShell|
Remove-PnPContentTypeFromList|PnP.PowerShell|[spo list contenttype remove](../cmd/spo/list/list-contenttype-remove.md)
Remove-PnPCustomAction|PnP.PowerShell|[spo customaction remove](../cmd/spo/customaction/customaction-remove.md)
Remove-PnPDeletedMicrosoft365Group|PnP.PowerShell|
Remove-PnPDeletedUnifiedGroup|PnP.PowerShell|
Remove-PnPEventReceiver|PnP.PowerShell|
Remove-PnPField|PnP.PowerShell|[spo field remove](../cmd/spo/field/field-remove.md)
Remove-PnPFieldFromContentType|PnP.PowerShell|[spo contenttype field remove](../cmd/spo/contenttype/contenttype-field-remove.md)
Remove-PnPFile|PnP.PowerShell|[spo file remove](../cmd/spo/file/file-remove.md)
Remove-PnPFileFromProvisioningTemplate|PnP.PowerShell|
Remove-PnPFileVersion|PnP.PowerShell|
Remove-PnPFolder|PnP.PowerShell|[spo folder remove](../cmd/spo/folder/folder-remove.md)
Remove-PnPGraphSubscription|PnP.PowerShell|
Remove-PnPGroup|PnP.PowerShell|[spo group remove](../cmd/spo/group/group-remove.md)
Remove-PnPHomeSite|PnP.PowerShell|[spo homesite remove](../cmd/spo/homesite/homesite-remove.md)
Remove-PnPHubSiteAssociation|PnP.PowerShell|[spo hubsite disconnect](../cmd/spo/hubsite/hubsite-disconnect.md)
Remove-PnPIndexedProperty|PnP.PowerShell|
Remove-PnPJavaScriptLink|PnP.PowerShell|
Remove-PnPKnowledgeHubSite|PnP.PowerShell|
Remove-PnPList|PnP.PowerShell|[spo list remove](../cmd/spo/list/list-remove.md)
Remove-PnPListItem|PnP.PowerShell|[spo listitem remove](../cmd/spo/listitem/listitem-remove.md)
Remove-PnPMicrosoft365Group|PnP.PowerShell|[aad o365group remove](../cmd/aad/o365group/o365group-remove.md)
Remove-PnPMicrosoft365GroupMember|PnP.PowerShell|[aad o365group user remove](../cmd/aad/o365group/o365group-user-remove.md)
Remove-PnPMicrosoft365GroupOwner|PnP.PowerShell|[aad o365group user remove](../cmd/aad/o365group/o365group-user-remove.md)
Remove-PnPNavigationNode|PnP.PowerShell|[spo navigation node remove](../cmd/spo/navigation/navigation-node-remove.md)
Remove-PnPOrgAssetsLibrary|PnP.PowerShell|[spo orgassetslibrary remove](../cmd/spo/orgassetslibrary/orgassetslibrary-remove.md)
Remove-PnPOrgNewsSite|PnP.PowerShell|[spo orgnewssite remove](../cmd/spo/orgnewssite/orgnewssite-remove.md)
Remove-PnPPlannerBucket|PnP.PowerShell|[planner bucket remove](../cmd/planner/bucket/bucket-remove.md)
Remove-PnPPlannerPlan|PnP.PowerShell|
Remove-PnPPlannerTask|PnP.PowerShell|
Remove-PnPPropertyBagValue|PnP.PowerShell|[spo propertybag remove](../cmd/spo/propertybag/propertybag-remove.md)
Remove-PnPPublishingImageRendition|PnP.PowerShell|
Remove-PnPRoleDefinition|PnP.PowerShell|
Remove-PnPSearchConfiguration|PnP.PowerShell|
Remove-PnPSiteClassification|PnP.PowerShell|
Remove-PnPSiteCollectionAdmin|PnP.PowerShell|
Remove-PnPSiteCollectionAppCatalog|PnP.PowerShell|[spo site appcatalog remove](../cmd/spo/site/site-appcatalog-remove.md)
Remove-PnPSiteDesign|PnP.PowerShell|[spo sitedesign remove](../cmd/spo/sitedesign/sitedesign-remove.md)
Remove-PnPSiteDesignTask|PnP.PowerShell|
Remove-PnPSiteScript|PnP.PowerShell|[spo sitescript remove](../cmd/spo/sitescript/sitescript-remove.md)
Remove-PnPStorageEntity|PnP.PowerShell|[spo storageentity remove](../cmd/spo/storageentity/storageentity-remove.md)
Remove-PnPStoredCredential|PnP.PowerShell|
Remove-PnPTaxonomyItem|PnP.PowerShell|
Remove-PnPTeamsApp|PnP.PowerShell|[teams app remove](../cmd/teams/app/app-remove.md)
Remove-PnPTeamsChannel|PnP.PowerShell|[teams channel remove](../cmd/teams/channel/channel-remove.md)
Remove-PnPTeamsChannelUser|PnP.PowerShell|[teams channel member remove](../cmd/teams/channel/channel-member-remove.md)
Remove-PnPTeamsTab|PnP.PowerShell|[teams tab remove](../cmd/teams/tab/tab-remove.md)
Remove-PnPTeamsTeam|PnP.PowerShell|[teams team remove](../cmd/teams/team/team-remove.md)
Remove-PnPTeamsUser|PnP.PowerShell|[teams user remove](../cmd/aad/o365group/o365group-user-remove.md)
Remove-PnPTenantCdnOrigin|PnP.PowerShell|[spo cdn origin remove](../cmd/spo/cdn/cdn-origin-remove.md)
Remove-PnPTenantSite|PnP.PowerShell|
Remove-PnPTenantTheme|PnP.PowerShell|[spo theme remove](../cmd/spo/theme/theme-remove.md)
Remove-PnPTermGroup|PnP.PowerShell|
Remove-PnPUnifiedGroup|PnP.PowerShell|[aad o365group remove](../cmd/aad/o365group/o365group-remove.md)
Remove-PnPUnifiedGroupMember|PnP.PowerShell|[aad o365group user remove](../cmd/aad/o365group/o365group-user-remove.md)
Remove-PnPUnifiedGroupOwner|PnP.PowerShell|[aad o365group user remove](../cmd/aad/o365group/o365group-user-remove.md)
Remove-PnPUser|PnP.PowerShell|[spo user remove](../cmd/spo/user/user-remove.md)
Remove-PnPUserFromGroup|PnP.PowerShell|
Remove-PnPView|PnP.PowerShell|[spo list view remove](../cmd/spo/list/list-view-remove.md)
Remove-PnPWeb|PnP.PowerShell|[spo web remove](../cmd/spo/web/web-remove.md)
Remove-PnPWebhookSubscription|PnP.PowerShell|[spo list webhook remove](../cmd/spo/list/list-webhook-remove.md)
Remove-PnPWebPart|PnP.PowerShell|
Remove-PnPWikiPage|PnP.PowerShell|
Remove-PnPWorkflowDefinition|PnP.PowerShell|
Remove-PnPWorkflowSubscription|PnP.PowerShell|
Rename-PnPFile|PnP.PowerShell|
Rename-PnPFolder|PnP.PowerShell|[spo folder rename](../cmd/spo/folder/folder-rename.md)
Request-PnPAccessToken|PnP.PowerShell|[util accesstoken get](../cmd/util/accesstoken/accesstoken-get.md)
Request-PnPReIndexList|PnP.PowerShell|
Request-PnPReIndexWeb|PnP.PowerShell|[spo web reindex](../cmd/spo/web/web-reindex.md)
Reset-PnPFileVersion|PnP.PowerShell|
Reset-PnPLabel|PnP.PowerShell|
Reset-PnPMicrosoft365GroupExpiration|PnP.PowerShell|[aad o365group renew](../cmd/aad/o365group/o365group-renew.md)
Reset-PnPUnifiedGroupExpiration|PnP.PowerShell|[aad o365group renew](../cmd/aad/o365group/o365group-renew.md)
Reset-PnPUserOneDriveQuotaToDefault|PnP.PowerShell|
Resolve-PnPFolder|PnP.PowerShell|
Restore-PnPDeletedMicrosoft365Group|PnP.PowerShell|[aad o365group restore](../cmd/aad/o365group/o365group-recyclebinitem-restore.md)
Restore-PnPDeletedUnifiedGroup|PnP.PowerShell|[aad o365group restore](../cmd/aad/o365group/o365group-recyclebinitem-restore.md)
Restore-PnPFileVersion|PnP.PowerShell|
Restore-PnPRecycleBinItem|PnP.PowerShell|
Restore-PnPTenantRecycleBinItem|PnP.PowerShell|[spo tenant recyclebinitem restore](../cmd/spo/tenant/tenant-recyclebinitem-restore.md)
Resume-PnPWorkflowInstance|PnP.PowerShell|
Revoke-PnPHubSiteRights|PnP.PowerShell|[spo hubsite rights revoke](../cmd/spo/hubsite/hubsite-rights-revoke.md)
Revoke-PnPSiteDesignRights|PnP.PowerShell|[spo sitedesign rights revoke](../cmd/spo/sitedesign/sitedesign-rights-revoke.md)
Revoke-PnPTenantServicePrincipalPermission|PnP.PowerShell|[spo serviceprincipal grant revoke](../cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md)
Save-PnPClientSidePageConversionLog|PnP.PowerShell|
Save-PnPProvisioningHierarchy|PnP.PowerShell|
Save-PnPProvisioningTemplate|PnP.PowerShell|
Save-PnPTenantTemplate|PnP.PowerShell|
Send-PnPMail|PnP.PowerShell|[spo mail send](../cmd/spo/mail/mail-send.md)
Set-PnPApplicationCustomizer|PnP.PowerShell|
Set-PnPAppSideLoading|PnP.PowerShell|
Set-PnPAuditing|PnP.PowerShell|
Set-PnPAvailablePageLayouts|PnP.PowerShell|
Set-PnPClientSidePage|PnP.PowerShell|[spo page set](../cmd/spo/page/page-set.md), [spo page header set](../cmd/spo/page/page-header-set.md)
Set-PnPClientSideText|PnP.PowerShell|
Set-PnPClientSideWebPart|PnP.PowerShell|
Set-PnPContext|PnP.PowerShell|
Set-PnPDefaultColumnValues|PnP.PowerShell|
Set-PnPDefaultContentTypeToList|PnP.PowerShell|[spo list contenttype default set](../cmd/spo/list/list-contenttype-default-set.md)
Set-PnPDefaultPageLayout|PnP.PowerShell|
Set-PnPDocumentSetField|PnP.PowerShell|
Set-PnPField|PnP.PowerShell|[spo field set](../cmd/spo/field/field-set.md)
Set-PnPFileCheckedIn|PnP.PowerShell|[spo file checkin](../cmd/spo/file/file-checkin.md)
Set-PnPFileCheckedOut|PnP.PowerShell|[spo file checkout](../cmd/spo/file/file-checkout.md)
Set-PnPFolderPermission|PnP.PowerShell|
Set-PnPFooter|PnP.PowerShell|
Set-PnPGraphSubscription|PnP.PowerShell|
Set-PnPGroup|PnP.PowerShell|
Set-PnPGroupPermissions|PnP.PowerShell|
Set-PnPHideDefaultThemes|PnP.PowerShell|[spo hidedefaultthemes set](../cmd/spo/hidedefaultthemes/hidedefaultthemes-set.md)
Set-PnPHomePage|PnP.PowerShell|[spo web set](../cmd/spo/web/web-set.md)
Set-PnPHomeSite|PnP.PowerShell|[spo homesite set](../cmd/spo/homesite/homesite-set.md)
Set-PnPHubSite|PnP.PowerShell|[spo hubsite set](../cmd/spo/hubsite/hubsite-set.md)
Set-PnPIndexedProperties|PnP.PowerShell|
Set-PnPInPlaceRecordsManagement|PnP.PowerShell|[spo site inplacerecordsmanagement set](../cmd/spo/site/site-inplacerecordsmanagement-set.md)
Set-PnPKnowledgeHubSite|PnP.PowerShell|[spo knowledgehub set](../cmd/spo/knowledgehub/knowledgehub-set.md)
Set-PnPLabel|PnP.PowerShell|[spo list label set](../cmd/spo/list/list-label-set.md)
Set-PnPList|PnP.PowerShell|[spo list set](../cmd/spo/list/list-set.md)
Set-PnPListInformationRightsManagement|PnP.PowerShell|
Set-PnPListItem|PnP.PowerShell|[spo listitem set](../cmd/spo/listitem/listitem-set.md)
Set-PnPListItemAsRecord|PnP.PowerShell|[spo listitem record declare](../cmd/spo/listitem/listitem-record-declare.md)
Set-PnPListItemPermission|PnP.PowerShell|
Set-PnPListPermission|PnP.PowerShell|
Set-PnPListRecordDeclaration|PnP.PowerShell|
Set-PnPMasterPage|PnP.PowerShell|
Set-PnPMicrosoft365Group|PnP.PowerShell|[aad o365group set](../cmd/aad/o365group/o365group-set.md)
Set-PnPMinimalDownloadStrategy|PnP.PowerShell|
Set-PnPPlannerBucket|PnP.PowerShell|[planner bucket set](../cmd/planner/bucket/bucket-set.md)
Set-PnPPlannerPlan|PnP.PowerShell|
Set-PnPPlannerTask|PnP.PowerShell|[planner task set](../cmd/planner/task/task-set.md)
Set-PnPPropertyBagValue|PnP.PowerShell|[spo propertybag set](../cmd/spo/propertybag/propertybag-set.md)
Set-PnPProvisioningTemplateMetadata|PnP.PowerShell|
Set-PnPRequestAccessEmails|PnP.PowerShell|
Set-PnPSearchConfiguration|PnP.PowerShell|
Set-PnPSearchSettings|PnP.PowerShell|
Set-PnPSite|PnP.PowerShell|[spo site set](../cmd/spo/site/site-set.md)
Set-PnPSiteClosure|PnP.PowerShell|
Set-PnPSiteDesign|PnP.PowerShell|[spo sitedesign set](../cmd/spo/sitedesign/sitedesign-set.md)
Set-PnPSitePolicy|PnP.PowerShell|
Set-PnPSiteScript|PnP.PowerShell|[spo sitescript set](../cmd/spo/sitescript/sitescript-set.md)
Set-PnPStorageEntity|PnP.PowerShell|[spo storageentity set](../cmd/spo/storageentity/storageentity-set.md)
Set-PnPTaxonomyFieldValue|PnP.PowerShell|
Set-PnPTeamsChannel|PnP.PowerShell|[teams channel set](../cmd/teams/channel/channel-set.md)
Set-PnPTeamsChannelUser|PnP.PowerShell|[teams channel member set](../cmd/teams/channel/channel-member-set.md)
Set-PnPTeamsTab|PnP.PowerShell|
Set-PnPTeamsTeam|PnP.PowerShell|[teams team set](../cmd/teams/team/team-set.md)
Set-PnPTeamsTeamArchivedState|PnP.PowerShell|[teams team archive](../cmd/teams/team/team-archive.md), [teams team unarchive](../cmd/teams/team/team-unarchive.md)
Set-PnPTeamsTeamPicture|PnP.PowerShell|
Set-PnPTenant|PnP.PowerShell|[spo tenant settings set](../cmd/spo/tenant/tenant-settings-set.md)
Set-PnPTenantAppCatalogUrl|PnP.PowerShell|
Set-PnPTenantCdnEnabled|PnP.PowerShell|[spo cdn set](../cmd/spo/cdn/cdn-set.md)
Set-PnPTenantCdnPolicy|PnP.PowerShell|[spo cdn policy set](../cmd/spo/cdn/cdn-policy-set.md)
Set-PnPTenantSite|PnP.PowerShell|[spo site classic set](../cmd/spo/site/site-classic-set.md)
Set-PnPTenantSyncClientRestriction|PnP.PowerShell|
Set-PnPTheme|PnP.PowerShell|
Set-PnPTheme|PnP.PowerShell|[spo theme apply](../cmd/spo/theme/theme-apply.md)
Set-PnPTraceLog|PnP.PowerShell|
Set-PnPUnifiedGroup|PnP.PowerShell|[aad o365group set](../cmd/aad/o365group/o365group-set.md)
Set-PnPUserOneDriveQuota|PnP.PowerShell|
Set-PnPUserProfileProperty|PnP.PowerShell|[spo userprofile set](../cmd/spo/userprofile/userprofile-set.md)
Set-PnPView|PnP.PowerShell|[spo list view set](../cmd/spo/list/list-view-set.md)
Set-PnPWeb|PnP.PowerShell|[spo web set](../cmd/spo/web/web-set.md)
Set-PnPWebhookSubscription|PnP.PowerShell|[spo list webhook set](../cmd/spo/list/list-webhook-set.md)
Set-PnPWebPartProperty|PnP.PowerShell|
Set-PnPWebPermission|PnP.PowerShell|
Set-PnPWebTheme|PnP.PowerShell|
Set-PnPWikiPageContent|PnP.PowerShell|
Start-PnPWorkflowInstance|PnP.PowerShell|
Stop-PnPFlowRun|PnP.PowerShell|[flow run cancel](../cmd/flow/run/run-cancel.md)
Stop-PnPWorkflowInstance|PnP.PowerShell|
Submit-PnPSearchQuery|PnP.PowerShell|[spo search](../cmd/spo/spo-search.md)
Submit-PnPTeamsChannelMessage|PnP.PowerShell|
Sync-PnPAppToTeams|PnP.PowerShell|
Test-PnPListItemIsRecord|PnP.PowerShell|[spo listitem isrecord](../cmd/spo/listitem/listitem-isrecord.md)
Test-PnPOffice365GroupAliasIsUsed|PnP.PowerShell|
Test-PnPProvisioningHierarchy|PnP.PowerShell|
Test-PnPTenantTemplate|PnP.PowerShell|
Uninstall-PnPApp|PnP.PowerShell|[spo app uninstall](../cmd/spo/app/app-uninstall.md)
Uninstall-PnPAppInstance|PnP.PowerShell|
Uninstall-PnPSolution|PnP.PowerShell|
Unpublish-PnPApp|PnP.PowerShell|[spo app retract](../cmd/spo/app/app-retract.md)
Unregister-PnPHubSite|PnP.PowerShell|[spo hubsite unregister](../cmd/spo/hubsite/hubsite-unregister.md)
Update-PnPApp|PnP.PowerShell|[spo app upgrade](../cmd/spo/app/app-upgrade.md)
Update-PnPSiteClassification|PnP.PowerShell|[aad siteclassification set](../cmd/aad/siteclassification/siteclassification-set.md)
Update-PnPTeamsApp|PnP.PowerShell|[teams app update](../cmd/teams/app/app-update.md)
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
