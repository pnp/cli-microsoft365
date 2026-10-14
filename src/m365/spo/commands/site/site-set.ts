import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo, SpoOperation } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SharingCapabilities } from '../site/SharingCapabilities.js';
import { FlowsPolicy } from './FlowsPolicy.js';
import { setTimeout } from 'timers/promises';

interface CommandArgs {
  options: Options;
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    }).alias('u'),
  title: z.string().optional().alias('t'),
  description: z.string().optional().alias('d'),
  classification: z.string().optional(),
  disableFlows: z.boolean().optional(),
  socialBarOnSitePagesDisabled: z.boolean().optional(),
  isPublic: z.boolean().optional(),
  owners: z.string().optional(),
  shareByEmailEnabled: z.boolean().optional(),
  siteDesignId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID.`
    }).optional(),
  sharingCapability: z.enum(['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']).optional(),
  siteLogoUrl: z.string().optional(),
  siteThumbnailUrl: z.string().optional(),
  resourceQuota: z.coerce.number().optional(),
  resourceQuotaWarningLevel: z.coerce.number().optional(),
  storageQuota: z.coerce.number().optional(),
  storageQuotaWarningLevel: z.coerce.number().optional(),
  allowSelfServiceUpgrade: z.boolean().optional(),
  lockState: z.enum(['Unlock', 'NoAdditions', 'ReadOnly', 'NoAccess']).optional(),
  noScriptSite: z.boolean().optional(),
  wait: z.boolean().optional()
});

export declare type Options = z.infer<typeof options>;

class SpoSiteSetCommand extends SpoCommand {
  private groupId: string | undefined;
  private siteId: string | undefined;
  private spoAdminUrl?: string;
  private context?: FormDigestInfo;
  private tenantId?: string;

  public get name(): string {
    return commands.SITE_SET;
  }

  public get description(): string {
    return 'Updates properties of the specified site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(
        (opts: Options) =>
          opts.classification !== undefined ||
          opts.disableFlows !== undefined ||
          opts.socialBarOnSitePagesDisabled !== undefined ||
          opts.title !== undefined ||
          opts.description !== undefined ||
          opts.isPublic !== undefined ||
          opts.owners !== undefined ||
          opts.shareByEmailEnabled !== undefined ||
          opts.siteDesignId !== undefined ||
          opts.sharingCapability !== undefined ||
          opts.siteLogoUrl !== undefined ||
          opts.siteThumbnailUrl !== undefined ||
          opts.resourceQuota !== undefined ||
          opts.resourceQuotaWarningLevel !== undefined ||
          opts.storageQuota !== undefined ||
          opts.storageQuotaWarningLevel !== undefined ||
          opts.noScriptSite !== undefined ||
          opts.allowSelfServiceUpgrade !== undefined ||
          opts.lockState !== undefined,
        {
          error: 'Specify at least one property to update'
        }
      )
      .refine(
        (opts: Options) => !opts.resourceQuota || !opts.resourceQuotaWarningLevel || opts.resourceQuotaWarningLevel <= opts.resourceQuota,
        {
          message: 'resourceQuotaWarningLevel must not exceed the resourceQuota',
          path: ['resourceQuotaWarningLevel']
        }
      )
      .refine(
        (opts: Options) => !opts.storageQuota || !opts.storageQuotaWarningLevel || opts.storageQuotaWarningLevel <= opts.storageQuota,
        {
          message: 'storageQuotaWarningLevel must not exceed the storageQuota',
          path: ['storageQuotaWarningLevel']
        }
      );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.tenantId = await spo.getTenantId(logger, this.debug);
      this.spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      this.context = await spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);

      await this.loadSiteIds(args.options.url, logger);

      if (this.isGroupConnectedSite()) {
        await this.updateGroupConnectedSite(logger, args);
      }
      else {
        await this.updateSharePointOnlySite(logger, args);
      }

      const siteProps = await this.updateSiteProperties(logger, args);
      await this.waitForSiteUpdateCompletion(logger, args, siteProps);
      await this.applySiteDesign(logger, args);
      await this.setLogo(logger, args);
      await this.setThumbnail(logger, args);
      const lockState = await this.updateSiteLockState(logger, args);
      await this.waitForSiteUpdateCompletion(logger, args, lockState);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async setLogo(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.siteLogoUrl === 'undefined') {
      return Promise.resolve();
    }

    if (this.debug) {
      await logger.logToStderr(`Setting the site its logo...`);
    }

    const logoUrl = args.options.siteLogoUrl ? urlUtil.getUrlRelativePath(args.options.siteLogoUrl) : "";

    const requestOptions: any = {
      url: `${args.options.url}/_api/siteiconmanager/setsitelogo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        aspect: 1,
        relativeLogoUrl: logoUrl,
        type: 0
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async setThumbnail(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.siteThumbnailUrl === 'undefined') {
      return;
    }

    if (this.debug) {
      await logger.logToStderr(`Setting the site thumbnail...`);
    }

    const thumbnailUrl = args.options.siteThumbnailUrl ? urlUtil.getUrlRelativePath(args.options.siteThumbnailUrl) : "";

    const requestOptions: any = {
      url: `${args.options.url}/_api/siteiconmanager/setsitelogo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        aspect: 0,
        relativeLogoUrl: thumbnailUrl,
        type: 0
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async updateSharePointOnlySite(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.debug) {
      await logger.logToStderr('Site is not group connected');
    }

    if (typeof args.options.isPublic !== 'undefined') {
      throw `The isPublic option can't be set on a site that is not groupified`;
    }

    await this.updateSiteDescription(logger, args);
    await this.updateSiteOwners(logger, args);
  }

  private async waitForSiteUpdateCompletion(logger: Logger, args: CommandArgs, response: string | undefined): Promise<void> {
    if (!response) {
      return;
    }

    const json: ClientSvcResponse = JSON.parse(response);
    const responseContent: ClientSvcResponseContents = json[0];

    if (responseContent.ErrorInfo) {
      throw responseContent.ErrorInfo.ErrorMessage;
    }

    const operation: SpoOperation = json[json.length - 1];
    const isComplete: boolean = operation.IsComplete;

    if (!args.options.wait || isComplete) {
      return;
    }

    await setTimeout(operation.PollingInterval);

    await spo.waitUntilFinished({
      operationId: JSON.stringify(operation._ObjectIdentity_),
      siteUrl: this.spoAdminUrl as string,
      logger,
      currentContext: this.context as FormDigestInfo,
      debug: this.debug,
      verbose: this.verbose
    });
  }

  private async updateSiteOwners(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.owners) {
      return;
    }

    if (this.verbose) {
      await logger.logToStderr(`Updating site owners ${args.options.url}...`);
    }

    await Promise.all(args.options.owners!.split(',').map(o => {
      return this.setAdmin(args.options.url, o.trim());
    }));
  }

  private async setAdmin(siteUrl: string, principal: string): Promise<void> {
    const requestOptions: any = {
      url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': this.context!.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">${formatting.escapeXml(siteUrl)}</Parameter><Parameter Type="String">${formatting.escapeXml(principal)}</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const response = await request.post<string>(requestOptions);
    const json: ClientSvcResponse = JSON.parse(response);
    const responseContent: ClientSvcResponseContents = json[0];

    if (responseContent.ErrorInfo) {
      throw responseContent.ErrorInfo.ErrorMessage;
    }
  }

  private async updateSiteDescription(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.description) {
      return;
    }

    if (this.verbose) {
      await logger.logToStderr(`Setting site description ${args.options.url}...`);
    }

    const requestOptions: any = {
      url: `${args.options.url}/_api/web`,
      headers: {
        'IF-MATCH': '*',
        'Accept': 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'X-RequestDigest': this.context!.FormDigestValue,
        'X-HTTP-Method': 'MERGE'
      },
      data: {
        Description: args.options.description
      },
      json: true
    };

    await request.post(requestOptions);
  }

  private async updateSiteLockState(logger: Logger, args: CommandArgs): Promise<string | undefined> {
    if (!args.options.lockState) {
      return;
    }

    if (this.verbose) {
      await logger.logToStderr(`Setting site lock state ${args.options.url}...`);
    }

    const requestOptions: any = {
      url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': (this.context as FormDigestInfo).FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">${formatting.escapeXml(args.options.lockState)}</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    return request.post(requestOptions);
  }

  private async updateGroupConnectedSite(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.debug) {
      await logger.logToStderr(`Site attached to group ${this.groupId}`);
    }

    if (typeof args.options.title === 'undefined' &&
      typeof args.options.description === 'undefined' &&
      typeof args.options.isPublic === 'undefined' &&
      typeof args.options.owners === 'undefined') {
      return;
    }

    const promises: Promise<void>[] = [];

    if (typeof args.options.title !== 'undefined') {
      const requestOptions: any = {
        url: `${this.spoAdminUrl}/_api/SPOGroup/UpdateGroupPropertiesBySiteId`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;charset=utf-8',
          'X-RequestDigest': this.context!.FormDigestValue
        },
        data: {
          groupId: this.groupId,
          siteId: this.siteId,
          displayName: args.options.title
        },
        responseType: 'json'
      };

      promises.push(request.post(requestOptions));
    }

    if (typeof args.options.isPublic !== 'undefined') {
      promises.push(entraGroup.setGroup(this.groupId as string, (args.options.isPublic === false), undefined, undefined, logger, this.verbose));
    }

    if (args.options.description) {
      promises.push(this.setGroupifiedSiteDescription(args.options.description));
    }

    promises.push(this.setGroupifiedSiteOwners(logger, args));

    await Promise.all(promises);
  }

  private setGroupifiedSiteDescription(description: string): Promise<void> {
    const requestOptions: any = {
      url: `https://graph.microsoft.com/v1.0/groups/${this.groupId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      data: {
        description: description
      }
    };

    return request.patch(requestOptions);
  }

  private async setGroupifiedSiteOwners(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.owners === 'undefined') {
      return;
    }

    const owners: string[] = args.options.owners.split(',').map(o => o.trim());

    if (this.verbose) {
      await logger.logToStderr('Retrieving user information to set group owners...');
    }

    const requestOptions: any = {
      url: `https://graph.microsoft.com/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);

    if (response.value.length === 0) {
      return;
    }

    await Promise.all(response.value.map(user => {
      const requestOptions: any = {
        url: `${this.spoAdminUrl}/_api/SP.Directory.DirectorySession/Group('${this.groupId}')/Owners/Add(objectId='${user.id}', principalName='')`,
        headers: {
          'content-type': 'application/json;odata=verbose'
        }
      };

      return request.post(requestOptions);
    }));
  }

  private async updateSiteProperties(logger: Logger, args: CommandArgs): Promise<string | undefined> {
    const isGroupConnectedSite = this.isGroupConnectedSite();
    const sharedProperties: string[] = ['classification', 'disableFlows', 'socialBarOnSitePagesDisabled', 'shareByEmailEnabled', 'sharingCapability', 'noScriptSite'];
    const siteProperties: string[] = ['title', 'resourceQuota', 'resourceQuotaWarningLevel', 'storageQuota', 'storageQuotaWarningLevel', 'allowSelfServiceUpgrade'];
    let properties: string[] = sharedProperties;

    if (!isGroupConnectedSite) {
      properties = properties.concat(siteProperties);
    }

    let updatedProperties: boolean = false;
    for (let i: number = 0; i < properties.length; i++) {
      if (typeof (args.options as any)[properties[i]] !== 'undefined') {
        updatedProperties = true;
        break;
      }
    }

    if (!updatedProperties) {
      return;
    }

    this.context = await spo.ensureFormDigest(this.spoAdminUrl!, logger, this.context, this.debug);

    if (this.verbose) {
      await logger.logToStderr(`Updating site ${args.options.url} properties...`);
    }

    let propertyId = 27;
    const payload: string[] = [];
    const sitePropertiesPayload: string[] = [];

    if (!isGroupConnectedSite) {
      if (args.options.title) {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Title"><Parameter Type="String">${formatting.escapeXml(args.options.title)}</Parameter></SetProperty>`);
      }
      if (args.options.resourceQuota) {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="UserCodeMaximumLevel"><Parameter Type="Double">${args.options.resourceQuota}</Parameter></SetProperty>`);
      }
      if (args.options.resourceQuotaWarningLevel) {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="UserCodeWarningLevel"><Parameter Type="Double">${args.options.resourceQuotaWarningLevel}</Parameter></SetProperty>`);
      }
      if (args.options.storageQuota) {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="StorageMaximumLevel"><Parameter Type="Int64">${args.options.storageQuota}</Parameter></SetProperty>`);
      }
      if (args.options.storageQuotaWarningLevel) {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="StorageWarningLevel"><Parameter Type="Int64">${args.options.storageQuotaWarningLevel}</Parameter></SetProperty>`);
      }
      if (args.options.allowSelfServiceUpgrade !== undefined) {
        payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="AllowSelfServiceUpgrade"><Parameter Type="Boolean">${args.options.allowSelfServiceUpgrade}</Parameter></SetProperty>`);
      }
    }
    if (typeof args.options.classification === 'string') {
      sitePropertiesPayload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Classification"><Parameter Type="String">${formatting.escapeXml(args.options.classification)}</Parameter></SetProperty>`);
    }
    if (args.options.disableFlows !== undefined) {
      const flowsPolicy: FlowsPolicy = args.options.disableFlows ? FlowsPolicy.Disabled : FlowsPolicy.Enabled;
      payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Enum">${flowsPolicy}</Parameter></SetProperty>`);
    }
    if (typeof args.options.socialBarOnSitePagesDisabled !== 'undefined') {
      payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="SocialBarOnSitePagesDisabled"><Parameter Type="Boolean">${args.options.socialBarOnSitePagesDisabled}</Parameter></SetProperty>`);
    }
    if (args.options.shareByEmailEnabled !== undefined) {
      sitePropertiesPayload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">${args.options.shareByEmailEnabled}</Parameter></SetProperty>`);
    }
    if (args.options.sharingCapability) {
      const sharingCapability: SharingCapabilities = SharingCapabilities[(args.options.sharingCapability as keyof typeof SharingCapabilities)];
      payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">${sharingCapability}</Parameter></SetProperty>`);
    }
    if (args.options.noScriptSite !== undefined) {
      const noScriptSite: number = args.options.noScriptSite ? 2 : 1;
      payload.push(`<SetProperty Id="${propertyId}" ObjectPathId="5" Name="DenyAddAndCustomizePages"><Parameter Type="Enum">${noScriptSite}</Parameter></SetProperty>`);
    }

    let response;
    let sitePropertiesResponse;

    if (sitePropertiesPayload.length > 0) {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.url}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': this.context.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${sitePropertiesPayload.join('')}</Actions><ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="5" ParentId="1" Name="Site" /></ObjectPaths></Request>`
      };

      sitePropertiesResponse = await request.post<string>(requestOptions);
    }

    if (payload.length > 0) {
      const pos = (this.tenantId as string).indexOf('|') + 1;

      const requestOptions: CliRequestOptions = {
        url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': this.context.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${payload.join('')}<ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|${(this.tenantId as string).substring(pos, (this.tenantId as string).indexOf('&'))}&#xA;SiteProperties&#xA;${formatting.encodeQueryParameter(args.options.url)}" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`
      };

      response = await request.post<string>(requestOptions);
    }

    return response || sitePropertiesResponse;
  }

  private async applySiteDesign(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.siteDesignId === 'undefined') {
      return;
    }

    return spo.applySiteDesign(args.options.url, args.options.siteDesignId, logger, this.verbose);
  }

  private async loadSiteIds(siteUrl: string, logger: Logger): Promise<void> {
    if (this.debug) {
      await logger.logToStderr('Loading site IDs...');
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/site?$select=GroupId,Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const siteInfo = await request.get<{ GroupId: string; Id: string }>(requestOptions);
    this.groupId = siteInfo.GroupId;
    this.siteId = siteInfo.Id;

    if (this.debug) {
      await logger.logToStderr(`Retrieved site IDs. siteId: ${this.siteId}, groupId: ${this.groupId}`);
    }
  }

  private isGroupConnectedSite(): boolean {
    return this.groupId !== '00000000-0000-0000-0000-000000000000';
  }

}

export default new SpoSiteSetCommand();