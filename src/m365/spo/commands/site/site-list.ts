import { Logger } from '../../../../cli/Logger.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SiteProperties } from './SiteProperties.js';
import { SPOSitePropertiesEnumerable } from './SPOSitePropertiesEnumerable.js';

enum SiteListType {
  TeamSite = 'TeamSite',
  CommunicationSite = 'CommunicationSite',
  fullyArchived = 'fullyArchived',
  recentlyArchived = 'recentlyArchived',
  archived = 'archived'
}

const options = globalOptionsZod
  .extend({
    type: zod.alias('t', zod.coercedEnum(SiteListType)).optional(),
    webTemplate: z.string().optional(),
    filter: z.string().optional(),
    withOneDriveSites: z.boolean().optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteListCommand extends SpoCommand {
  private allSites?: SiteProperties[];

  public get name(): string {
    return commands.SITE_LIST;
  }

  public get description(): string {
    return 'Lists sites of the given type';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(o => !(o.type && o.webTemplate), {
        message: 'Specify either type or webTemplate, but not both'
      })
      .refine(o => !(o.withOneDriveSites && (o.type !== undefined || o.webTemplate !== undefined)), {
        message: 'When using withOneDriveSites, don\'t specify the type or webTemplate options'
      });
  }

  public alias(): string[] | undefined {
    return [commands.SITE_LIST];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const webTemplate: string = this.getWebTemplateId(args.options);
    const includeOneDriveSites: boolean = args.options.withOneDriveSites || false;
    const personalSite: string = includeOneDriveSites === false ? '0' : '1';
    const effectiveFilter: string = this.buildFilter(args.options);

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of site collections...`);
      }

      this.allSites = [];

      await this.getAllSites(spoAdminUrl, formatting.escapeXml(effectiveFilter), '0', personalSite, webTemplate, undefined, logger);
      await logger.log(this.allSites);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAllSites(spoAdminUrl: string, filter: string | undefined, startIndex: string | undefined, personalSite: string, webTemplate: string, formDigest: FormDigestInfo | undefined | undefined, logger: Logger): Promise<void> {
    const res: FormDigestInfo = await spo.ensureFormDigest(spoAdminUrl, logger, formDigest, this.debug);

    const requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">${filter}</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">${personalSite}</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">${webTemplate}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': res.FormDigestValue
      },
      data: requestBody
    };

    const response: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(response);
    const responseContent: ClientSvcResponseContents = json[0];

    if (responseContent.ErrorInfo) {
      throw responseContent.ErrorInfo.ErrorMessage;
    }
    else {
      const sites: SPOSitePropertiesEnumerable = json[json.length - 1];
      this.allSites!.push(...sites._Child_Items_);

      if (sites.NextStartIndexFromSharePoint) {
        await this.getAllSites(spoAdminUrl, filter, sites.NextStartIndexFromSharePoint, personalSite, webTemplate, formDigest, logger);
      }

      return;
    }
  }

  private getWebTemplateId(options: Options): string {
    if (options.webTemplate) {
      return options.webTemplate;
    }

    if (options.withOneDriveSites) {
      return '';
    }

    switch (options.type) {
      case "TeamSite":
        return 'GROUP#0';
      case "CommunicationSite":
        return 'SITEPAGEPUBLISHING#0';
      default:
        return '';
    }
  }

  private buildFilter(options: Options): string {
    const providedFilter: string = options.filter || '';

    let archivedFilter: string = '';
    switch (options.type) {
      case 'fullyArchived':
        archivedFilter = "ArchiveStatus -eq 'FullyArchived'";
        break;
      case 'recentlyArchived':
        archivedFilter = "ArchiveStatus -eq 'RecentlyArchived'";
        break;
      case 'archived':
        archivedFilter = "ArchiveStatus -ne 'NotArchived'";
        break;
    }

    if (archivedFilter && providedFilter) {
      return `${providedFilter} and ${archivedFilter}`;
    }

    return archivedFilter || providedFilter;
  }
}

export default new SpoSiteListCommand();