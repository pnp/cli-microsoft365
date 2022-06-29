import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, formatting, FormDigestInfo, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteProperties } from './SiteProperties';
import { SPOSitePropertiesEnumerable } from './SPOSitePropertiesEnumerable';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  webTemplate?: string;
  filter?: string;
  deleted?: boolean;
  includeOneDriveSites?: boolean;
}

class SpoSiteListCommand extends SpoCommand {
  private allSites?: SiteProperties[];

  public get name(): string {
    return commands.SITE_LIST;
  }

  public get description(): string {
    return 'Lists sites of the given type';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.webTemplate = args.options.webTemplate;
    telemetryProps.filter = (!(!args.options.filter)).toString();
    telemetryProps.includeOneDriveSites = typeof args.options.includeOneDriveSites !== 'undefined';
    telemetryProps.deleted = typeof args.options.deleted !== 'undefined';
    telemetryProps.siteType = args.options.type || 'TeamSite';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const webTemplate: string = this.getWebTemplateId(args.options);
    const includeOneDriveSites: boolean = args.options.includeOneDriveSites || false;
    const personalSite: string = includeOneDriveSites === false ? '0' : '1';
    let spoAdminUrl: string = '';

    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<void> => {
        spoAdminUrl = _spoAdminUrl;

        if (this.verbose) {
          logger.logToStderr(`Retrieving list of site collections...`);
        }

        this.allSites = [];

        return this.getAllSites(spoAdminUrl, formatting.escapeXml(args.options.filter || ''), '0', personalSite, webTemplate, undefined, args.options.deleted, logger);
      })
      .then(_ => {
        logger.log(this.allSites);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private getAllSites(spoAdminUrl: string, filter: string | undefined, startIndex: string | undefined, personalSite: string, webTemplate: string, formDigest: FormDigestInfo | undefined, deleted: boolean | undefined, logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      spo
        .ensureFormDigest(spoAdminUrl, logger, formDigest, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          let requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">${filter}</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">${personalSite}</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">${webTemplate}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
          if (deleted) {
            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties><Property Name="NextStartIndexFromSharePoint" ScalarProperty="true" /></Properties></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="GetDeletedSitePropertiesFromSharePoint"><Parameters><Parameter Type="String">${startIndex}</Parameter></Parameters></Method></ObjectPaths></Request>`;
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            data: requestBody
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            reject(response.ErrorInfo.ErrorMessage);
            return;
          }
          else {
            const sites: SPOSitePropertiesEnumerable = json[json.length - 1];
            this.allSites!.push(...sites._Child_Items_);

            if (sites.NextStartIndexFromSharePoint) {
              this
                .getAllSites(spoAdminUrl, filter, sites.NextStartIndexFromSharePoint, personalSite, webTemplate, formDigest, deleted, logger)
                .then(_ => resolve(), err => reject(err));
            }
            else {
              resolve();
            }
          }
        }, err => reject(err));
    });
  }

  /* 
    The type property currently defaults to Teamsite. 
    It makes more sense to default to All. Certainly after adding the 'includeOneDriveSites' option.
    Changing this will be a breaking change. We'll remove the default the next major version.
  */
  private getWebTemplateId(options: Options): string {
    if (options.webTemplate) {
      return options.webTemplate;
    }

    if (options.includeOneDriveSites) {
      return '';
    }

    let siteType = options.type;

    if (!siteType) {
      siteType = 'TeamSite';
    }

    switch (siteType) {
      case "TeamSite":
        return 'GROUP#0';
      case "CommunicationSite":
        return 'SITEPAGEPUBLISHING#0';
      default:
        return '';
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --type [type]',
        // To not introduce a breaking change, 'All' has been added.
        // You should use all when using '--includeOneDriveSites'
        autocomplete: ['TeamSite', 'CommunicationSite', 'All']
      },
      {
        option: '--webTemplate [webTemplate]'
      },
      {
        option: '-f, --filter [filter]'
      },
      {
        option: '--includeOneDriveSites'
      },
      {
        option: '--deleted'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: any): string | boolean {
    if (args.options.type && args.options.webTemplate) {
      return 'Specify either type or webTemplate, but not both';
    }

    const typeValues = ['TeamSite', 'CommunicationSite', 'All'];
    if (args.options.type &&
      typeValues.indexOf(args.options.type) < 0) {
      return `${args.options.type} is not a valid value for the type option. Allowed values are ${typeValues.join('|')}`;
    }

    if (args.options.includeOneDriveSites
      && (!args.options.type || args.options.type !== 'All')) {
      return 'When using includeOneDriveSites, specify All as value for type';
    }

    return true;
  }
}

module.exports = new SpoSiteListCommand();