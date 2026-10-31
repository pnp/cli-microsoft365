import { setTimeout } from 'timers/promises';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo, SpoOperation } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { DeletedSiteProperties } from './DeletedSiteProperties.js';
import { SiteProperties } from './SiteProperties.js';
import { brandCenter } from '../../../../utils/brandCenter.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: z.enum(['TeamSite', 'CommunicationSite', 'ClassicSite', 'BrandCenter']).optional(),
  title: z.string().optional().alias('t'),
  alias: z.string().optional().alias('a'),
  url: z.string().optional().alias('u'),
  timeZone: z.coerce.number().optional().alias('z'),
  description: z.string().optional().alias('d'),
  lcid: z.coerce.number().optional().alias('l'),
  owners: z.string().optional(),
  isPublic: z.boolean().optional(),
  classification: z.string().optional().alias('c'),
  siteDesign: z.enum(['Topic', 'Showcase', 'Blank']).optional(),
  siteDesignId: z.string().optional(),
  shareByEmailEnabled: z.boolean().optional(),
  webTemplate: z.string().optional().alias('w'),
  resourceQuota: z.coerce.number().optional(),
  resourceQuotaWarningLevel: z.coerce.number().optional(),
  storageQuota: z.coerce.number().optional(),
  storageQuotaWarningLevel: z.coerce.number().optional(),
  removeDeletedSite: z.boolean().optional(),
  withAppCatalog: z.boolean().optional(),
  wait: z.boolean().optional(),
  force: z.boolean().optional()
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface CreateGroupExResponse {
  DocumentsUrl: string;
  ErrorMessage: string;
  GroupId: string;
  SiteStatus: number;
  SiteUrl: string;
}

class SpoSiteAddCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private spoAdminUrl?: string;

  private get supportedLcids(): number[] {
    // Languages supported by SharePoint
    // https://support.microsoft.com/en-us/office/languages-supported-by-sharepoint-dfbf3652-2902-4809-be21-9080b6512fff
    // https://github.com/pnp/PnP-PowerShell/wiki/Supported-LCIDs-by-SharePoint
    return [1025, 1068, 1069, 5146, 1026, 1027, 2052, 1028, 1050, 1029, 1030, 1043, 1033, 1061, 1035, 1036, 1110, 1031, 1032, 1037, 1081, 1038, 1057, 2108, 1040, 1041, 1087, 1042, 1062, 1063, 1071, 1086, 1044, 1045, 1046, 2070, 1048, 1049, 10266, 2074, 1051, 1060, 3082, 1053, 1054, 1055, 1058, 1066, 1106];
  }

  public get name(): string {
    return commands.SITE_ADD;
  }

  public get description(): string {
    return 'Creates new SharePoint Online site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    const isTeamSite = (o: Options): boolean => o.type === undefined || o.type === 'TeamSite';
    const isCommunicationSite = (o: Options): boolean => o.type === 'CommunicationSite' || o.type === 'BrandCenter';
    const isClassicSite = (o: Options): boolean => o.type === 'ClassicSite';

    return schema
      .refine(o => !isTeamSite(o) || o.alias !== undefined, {
        error: 'Required option alias missing',
        path: ['alias'],
        params: {
          customCode: 'required'
        }
      })
      .refine(o => !isTeamSite(o) || !(o.url || o.siteDesign || o.removeDeletedSite || o.wait || o.shareByEmailEnabled || o.siteDesignId || o.timeZone || o.resourceQuota || o.resourceQuotaWarningLevel || o.storageQuota || o.storageQuotaWarningLevel || o.webTemplate || o.force), {
        error: 'Type TeamSite supports only the parameters title, lcid, alias, owners, classification, isPublic, and description'
      })
      .refine(o => !isCommunicationSite(o) || o.url !== undefined, {
        error: 'Required option url missing',
        path: ['url'],
        params: {
          customCode: 'required'
        }
      })
      .refine(o => !isCommunicationSite(o) || o.url === undefined || validation.isValidSharePointUrl(o.url) === true, {
        error: e => `'${(e.input as Options).url}' is not a valid SharePoint Online site URL.`,
        path: ['url']
      })
      .refine(o => !isCommunicationSite(o) || o.owners === undefined || o.owners.indexOf(',') === -1, {
        error: 'The CommunicationSite supports only one owner in the owners option',
        path: ['owners']
      })
      .refine(o => !isCommunicationSite(o) || o.siteDesignId === undefined || validation.isValidGuid(o.siteDesignId) === true, {
        error: e => `${(e.input as Options).siteDesignId} is not a valid GUID`,
        path: ['siteDesignId']
      })
      .refine(o => !isCommunicationSite(o) || !(o.siteDesign && o.siteDesignId), {
        error: 'Specify siteDesign or siteDesignId but not both'
      })
      .refine(o => !isCommunicationSite(o) || !(o.timeZone || o.isPublic || o.removeDeletedSite || o.wait || o.alias || o.resourceQuota || o.resourceQuotaWarningLevel || o.storageQuota || o.storageQuotaWarningLevel || o.webTemplate), {
        error: 'Type CommunicationSite supports only the parameters url, title, lcid, classification, siteDesign, shareByEmailEnabled, siteDesignId, owners, description, and force'
      })
      .refine(o => !isClassicSite(o) || o.url !== undefined, {
        error: 'Required option url missing',
        path: ['url'],
        params: {
          customCode: 'required'
        }
      })
      .refine(o => !isClassicSite(o) || o.url === undefined || validation.isValidSharePointUrl(o.url) === true, {
        error: e => `'${(e.input as Options).url}' is not a valid SharePoint Online site URL.`,
        path: ['url']
      })
      .refine(o => !isClassicSite(o) || o.owners !== undefined, {
        error: 'Required option owner missing',
        path: ['owners'],
        params: {
          customCode: 'required'
        }
      })
      .refine(o => !isClassicSite(o) || o.owners === undefined || o.owners.indexOf(',') === -1, {
        error: 'The ClassicSite supports only one owner in the owners options',
        path: ['owners']
      })
      .refine(o => !isClassicSite(o) || o.timeZone !== undefined, {
        error: 'Required option timeZone missing',
        path: ['timeZone'],
        params: {
          customCode: 'required'
        }
      })
      .refine(o => !isClassicSite(o) || o.resourceQuotaWarningLevel === undefined || o.resourceQuota !== undefined, {
        error: 'You cannot specify resourceQuotaWarningLevel without specifying resourceQuota',
        path: ['resourceQuotaWarningLevel']
      })
      .refine(o => !isClassicSite(o) || o.resourceQuotaWarningLevel === undefined || o.resourceQuota === undefined || o.resourceQuotaWarningLevel <= o.resourceQuota, {
        error: 'resourceQuotaWarningLevel cannot exceed resourceQuota',
        path: ['resourceQuotaWarningLevel']
      })
      .refine(o => !isClassicSite(o) || o.storageQuotaWarningLevel === undefined || o.storageQuota !== undefined, {
        error: 'You cannot specify storageQuotaWarningLevel without specifying storageQuota',
        path: ['storageQuotaWarningLevel']
      })
      .refine(o => !isClassicSite(o) || o.storageQuotaWarningLevel === undefined || o.storageQuota === undefined || o.storageQuotaWarningLevel <= o.storageQuota, {
        error: 'storageQuotaWarningLevel cannot exceed storageQuota',
        path: ['storageQuotaWarningLevel']
      })
      .refine(o => !isClassicSite(o) || !(o.classification || o.shareByEmailEnabled || o.siteDesignId || o.alias || o.isPublic || o.force), {
        error: 'Type ClassicSite supports only the parameters url, title, lcid, storageQuota, storageQuotaWarningLevel, resourceQuota, resourceQuotaWarningLevel, webTemplate, owners, and description'
      })
      .refine(o => o.lcid === undefined || o.lcid >= 0, {
        error: e => `LCID must be greater than 0 (${(e.input as Options).lcid})`,
        path: ['lcid']
      })
      .refine(o => o.lcid === undefined || this.supportedLcids.indexOf(o.lcid) >= 0, {
        error: e => `LCID ${(e.input as Options).lcid} is not valid. See https://support.microsoft.com/en-us/office/languages-supported-by-sharepoint-dfbf3652-2902-4809-be21-9080b6512fff for the languages supported by SharePoint.`,
        path: ['lcid']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isClassicSite: boolean = args.options.type === 'ClassicSite';

    const siteUrl = isClassicSite
      ? await this.createClassicSite(logger, args)
      : await this.createModernSite(logger, args);

    if (siteUrl && args.options.withAppCatalog) {
      await this.addAppCatalog(siteUrl, logger);
    }

    await logger.log(siteUrl);
  }

  private async createModernSite(logger: Logger, args: CommandArgs): Promise<string | undefined> {
    const isTeamSite: boolean = args.options.type !== 'CommunicationSite' && args.options.type !== 'BrandCenter';

    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Creating new site...`);
      }

      let requestOptions: any = {};

      if (isTeamSite) {
        requestOptions = {
          url: `${spoUrl}/_api/GroupSiteManager/CreateGroupEx`,
          headers: {
            'content-type': 'application/json; odata=verbose; charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            displayName: args.options.title,
            alias: args.options.alias,
            isPublic: args.options.isPublic,
            optionalParams: {
              Description: args.options.description || '',
              CreationOptions: {
                results: [],
                Classification: args.options.classification || ''
              }
            }
          }
        };

        if (args.options.lcid) {
          requestOptions.data.optionalParams.CreationOptions.results.push(`SPSiteLanguage:${args.options.lcid}`);
        }

        if (args.options.owners) {
          requestOptions.data.optionalParams.Owners = {
            results: args.options.owners.split(',').map(o => o.trim())
          };
        }
      }
      else {
        let siteDesignId: string = '';
        if (args.options.siteDesignId) {
          siteDesignId = args.options.siteDesignId;
        }
        else {
          if (args.options.siteDesign) {
            switch (args.options.siteDesign) {
              case 'Topic':
                siteDesignId = '00000000-0000-0000-0000-000000000000';
                break;
              case 'Showcase':
                siteDesignId = '6142d2a0-63a5-4ba0-aede-d9fefca2c767';
                break;
              case 'Blank':
                siteDesignId = 'f6cc5403-0d63-442e-96c0-285923709ffc';
                break;
            }
          }
          else {
            siteDesignId = '00000000-0000-0000-0000-000000000000';
          }
        }

        requestOptions = {
          url: `${spoUrl}/_api/SPSiteManager/Create`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            request: {
              Title: args.options.title,
              Url: args.options.url,
              ShareByEmailEnabled: args.options.shareByEmailEnabled,
              Description: args.options.description || '',
              Classification: args.options.classification || '',
              WebTemplate: 'SITEPAGEPUBLISHING#0',
              SiteDesignId: siteDesignId
            }
          }
        };

        if (args.options.lcid) {
          requestOptions.data.request.Lcid = args.options.lcid;
        }

        if (args.options.owners) {
          requestOptions.data.request.Owner = args.options.owners;
        }

        if (args.options.type === 'BrandCenter') {
          await this.addBrandCenter(requestOptions.data.request, logger, args.options.force || false);
        }
      }

      const response = await request.post<CreateGroupExResponse>(requestOptions);

      if (isTeamSite) {
        if (response.ErrorMessage !== null) {
          throw response.ErrorMessage;
        }

        return response.SiteUrl;
      }
      else {
        if (response.SiteStatus !== 2) {
          throw 'An error has occurred while creating the site';
        }

        return response.SiteUrl;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);

      return;
    }
  }

  public async createClassicSite(logger: Logger, args: CommandArgs): Promise<string | undefined> {
    try {
      this.spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      this.context = await spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);

      let exists: boolean;
      if (args.options.removeDeletedSite) {
        exists = await this.siteExists(args.options.url as string, logger);
      }
      else {
        // assume site doesn't exist
        exists = false;
      }

      if (exists) {
        if (this.verbose) {
          await logger.logToStderr('Site exists in the recycle bin');
        }

        await this.deleteSiteFromTheRecycleBin(args.options.url as string, args.options.wait, logger);
      }
      else {
        if (this.verbose) {
          await logger.logToStderr('Site not found');
        }
      }

      this.context = await spo.ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Creating site collection ${args.options.url}...`);
      }

      const lcid: number = typeof args.options.lcid === 'number' ? args.options.lcid : 1033;
      const storageQuota: number = typeof args.options.storageQuota === 'number' ? args.options.storageQuota : 100;
      const storageQuotaWarningLevel: number = typeof args.options.storageQuotaWarningLevel === 'number' ? args.options.storageQuotaWarningLevel : 100;
      const resourceQuota: number = typeof args.options.resourceQuota === 'number' ? args.options.resourceQuota : 0;
      const resourceQuotaWarningLevel: number = typeof args.options.resourceQuotaWarningLevel === 'number' ? args.options.resourceQuotaWarningLevel : 0;
      const webTemplate: string = args.options.webTemplate || 'STS#0';

      const requestOptions: CliRequestOptions = {
        url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': this.context.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">${lcid}</Property><Property Name="Owner" Type="String">${formatting.escapeXml(args.options.owners)}</Property><Property Name="StorageMaximumLevel" Type="Int64">${storageQuota}</Property><Property Name="StorageWarningLevel" Type="Int64">${storageQuotaWarningLevel}</Property><Property Name="Template" Type="String">${formatting.escapeXml(webTemplate)}</Property><Property Name="TimeZoneId" Type="Int32">${args.options.timeZone}</Property><Property Name="Title" Type="String">${formatting.escapeXml(args.options.title)}</Property><Property Name="Url" Type="String">${formatting.escapeXml(args.options.url)}</Property><Property Name="UserCodeMaximumLevel" Type="Double">${resourceQuota}</Property><Property Name="UserCodeWarningLevel" Type="Double">${resourceQuotaWarningLevel}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      const response = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(response);
      const responseContent: ClientSvcResponseContents = json[0];

      if (responseContent.ErrorInfo) {
        throw responseContent.ErrorInfo.ErrorMessage;
      }

      const operation: SpoOperation = json[json.length - 1];
      const isComplete: boolean = operation.IsComplete;

      if ((!args.options.wait && !args.options.withAppCatalog) || isComplete) {
        return args.options.url;
      }

      await setTimeout(operation.PollingInterval);
      await spo.waitUntilFinished({
        operationId: JSON.stringify(operation._ObjectIdentity_),
        siteUrl: this.spoAdminUrl as string,
        logger,
        currentContext: this.context as FormDigestInfo,
        verbose: this.verbose,
        debug: this.debug
      });

      return args.options.url;
    }
    catch (err: any) {
      this.handleRejectedPromise(err);

      return;
    }
  }

  private async siteExists(url: string, logger: Logger): Promise<boolean> {
    this.context = await spo.ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug);

    if (this.verbose) {
      await logger.logToStderr(`Checking if the site ${url} exists...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': this.context.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`
    };

    const response: any = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(response);
    const responseContent: ClientSvcResponseContents = json[0];

    if (responseContent.ErrorInfo) {
      if (responseContent.ErrorInfo.ErrorTypeName === 'Microsoft.Online.SharePoint.Common.SpoNoSiteException') {
        return await this.siteExistsInTheRecycleBin(url, logger);
      }

      throw responseContent.ErrorInfo.ErrorMessage;
    }
    else {
      const site: SiteProperties = json[json.length - 1];

      return site.Status === 'Recycled';
    }
  }

  private async siteExistsInTheRecycleBin(url: string, logger: Logger): Promise<boolean> {
    if (this.verbose) {
      await logger.logToStderr(`Site doesn't exist. Checking if the site ${url} exists in the recycle bin...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': (this.context as FormDigestInfo).FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const res: any = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      if (response.ErrorInfo.ErrorTypeName === 'Microsoft.SharePoint.Client.UnknownError') {
        return false;
      }

      throw response.ErrorInfo.ErrorMessage;
    }

    const site: DeletedSiteProperties = json[json.length - 1];

    return site.Status === 'Recycled';
  }

  private async deleteSiteFromTheRecycleBin(url: string, wait: boolean | undefined, logger: Logger): Promise<void> {
    this.context = await spo.ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug);

    if (this.verbose) {
      await logger.logToStderr(`Deleting site ${url} from the recycle bin...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': this.context.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const response: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(response);
    const responseContent: ClientSvcResponseContents = json[0];

    if (responseContent.ErrorInfo) {
      throw responseContent.ErrorInfo.ErrorMessage;
    }

    const operation: SpoOperation = json[json.length - 1];
    const isComplete: boolean = operation.IsComplete;

    if (!wait || isComplete) {
      return;
    }

    await setTimeout(operation.PollingInterval);
    await spo.waitUntilFinished({
      operationId: JSON.stringify(operation._ObjectIdentity_),
      siteUrl: this.spoAdminUrl as string,
      logger,
      currentContext: this.context as FormDigestInfo,
      verbose: this.verbose,
      debug: this.debug
    });
  }

  private async addAppCatalog(url: string, logger: Logger): Promise<void> {
    try {
      this.spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      this.context = await spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Adding site collection app catalog...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': this.context.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="38" ObjectPathId="37" /><ObjectPath Id="40" ObjectPathId="39" /><ObjectPath Id="42" ObjectPathId="41" /><ObjectPath Id="44" ObjectPathId="43" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Constructor Id="37" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="39" ParentId="37" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Property Id="41" ParentId="39" Name="RootWeb" /><Property Id="43" ParentId="41" Name="TenantAppCatalog" /><Property Id="45" ParentId="43" Name="SiteCollectionAppCatalogsSites" /><Method Id="47" ParentId="45" Name="Add"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      const response: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(response);
      const responseContents: ClientSvcResponseContents = json[0];

      if (responseContents.ErrorInfo) {
        throw responseContents.ErrorInfo.ErrorMessage;
      }

      if (this.verbose) {
        await logger.logToStderr('Site collection app catalog created');
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async addBrandCenter(requestData: any, logger: Logger, force: boolean): Promise<void> {
    const brandingCenterConfiguration = await brandCenter.getBrandCenterConfiguration(logger, this.debug);

    if (brandingCenterConfiguration.IsBrandCenterSiteFeatureEnabled) {
      throw Error('Brand center site is already created in the tenant.');
    }

    const warningMessage = `You agree to activate this site as your official brand center site and turn on the brand center app for use in your organization. Storage locations will be created for uploading files to brand center and managing them. Any uploaded files will be stored in the cloud and managed in a public content delivery network (CDN). The files will be accessible to anyone who is able to extract the URLs that point to them.
Don't use this feature if your files contain proprietary information, or if you don't have the necessary cloud hosting rights to use them. After creation, that site cannot be deleted.`;

    if (force) {
      await logger.logToStderr(warningMessage);
    }
    else {
      const result = await cli.promptForConfirmation({
        message: `${warningMessage}\n\nDo you want to proceed?`
      });

      if (!result) {
        throw Error('Operation cancelled by the user.');
      }
    }

    const brandCenterFeatureId = '99cd6e8b-189b-4611-ae89-f89105876e43';
    requestData.AdditionalSiteFeatureIds = [brandCenterFeatureId];
  }
}

export default new SpoSiteAddCommand();