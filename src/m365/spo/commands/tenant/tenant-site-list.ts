import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { TenantSiteProperties } from './TenantSiteProperties.js';
import { SPOTenantSitePropertiesEnumerable } from './SPOTenantSitePropertiesEnumerable.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  webTemplate?: string;
  filter?: string;
  includeOneDriveSites?: boolean;
  withOneDriveSites?: boolean;
}

class SpoTenantSiteListCommand extends SpoCommand {
  private allSites?: TenantSiteProperties[];

  public get name(): string {
    return commands.TENANT_SITE_LIST;
  }

  public get description(): string {
    return 'Lists sites of the given type';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        webTemplate: args.options.webTemplate,
        type: args.options.type,
        filter: (!(!args.options.filter)).toString(),
        includeOneDriveSites: typeof args.options.includeOneDriveSites !== 'undefined',
        withOneDriveSites: typeof args.options.withOneDriveSites !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '--webTemplate [webTemplate]'
      },
      {
        option: '--filter [filter]'
      },
      {
        option: '--includeOneDriveSites'
      },
      {
        option: '--withOneDriveSites'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type && args.options.webTemplate) {
          return 'Specify either type or webTemplate, but not both';
        }

        const typeValues = ['TeamSite', 'CommunicationSite'];
        if (args.options.type &&
          typeValues.indexOf(args.options.type) < 0) {
          return `${args.options.type} is not a valid value for the type option. Allowed values are ${typeValues.join('|')}`;
        }

        if (args.options.includeOneDriveSites
          && (args.options.type || args.options.webTemplate)) {
          return 'When using includeOneDriveSites, don\'t specify the type or webTemplate options';
        }

        if (args.options.withOneDriveSites
          && (args.options.type || args.options.webTemplate)) {
          return 'When using withOneDriveSites, don\'t specify the type or webTemplate options';
        }

        return true;
      }
    );
  }

  public alias(): string[] | undefined {
    return [commands.SITE_LIST];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.includeOneDriveSites) {
      await this.warn(logger, `Parameter 'includeOneDriveSites' is deprecated. Please use 'withOneDriveSites' instead`);
    }

    const webTemplate: string = this.getWebTemplateId(args.options);
    const includeOneDriveSites: boolean = (args.options.includeOneDriveSites || args.options.withOneDriveSites) || false;
    const personalSite: string = includeOneDriveSites === false ? '0' : '1';

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of site collections...`);
      }

      this.allSites = [];

      await this.getAllSites(spoAdminUrl, formatting.escapeXml(args.options.filter || ''), '0', personalSite, webTemplate, undefined, logger);
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
      const sites: SPOTenantSitePropertiesEnumerable = json[json.length - 1];
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

    if (options.includeOneDriveSites) {
      return '';
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
}

export default new SpoTenantSiteListCommand();