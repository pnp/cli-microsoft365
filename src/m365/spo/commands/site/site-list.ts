import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo } from '../../spo';
import { SiteProperties } from './SiteProperties';
import { SPOSitePropertiesEnumerable } from './SPOSitePropertiesEnumerable';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  filter?: string;
  deleted?: boolean;
}

class SpoSiteListCommand extends SpoCommand {
  private allSites?: SiteProperties[];

  public get name(): string {
    return commands.SITE_LIST;
  }

  public get description(): string {
    return 'Lists modern sites of the given type';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.siteType = args.options.type || 'TeamSite';
    telemetryProps.filter = (!(!args.options.filter)).toString();
    telemetryProps.deleted = args.options.deleted;
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const siteType: string = args.options.type || 'TeamSite';
    const webTemplate: string = siteType === 'TeamSite' ? 'GROUP#0' : 'SITEPAGEPUBLISHING#0';
    let spoAdminUrl: string;

    this
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<void> => {
        spoAdminUrl = _spoAdminUrl;

        if (this.verbose) {
          logger.logToStderr(`Retrieving list of site collections...`);
        }

        this.allSites = [];

        return this.getAllSites(spoAdminUrl, Utils.escapeXml(args.options.filter || ''), '0', webTemplate, undefined, args.options.deleted, logger);
      })
      .then(_ => {
        logger.log(this.allSites);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private getAllSites(spoAdminUrl: string, filter: string | undefined, startIndex: string | undefined, webTemplate: string, formDigest: FormDigestInfo | undefined, deleted: boolean | undefined, logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .ensureFormDigest(spoAdminUrl, logger, formDigest, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          let requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">${filter}</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">${webTemplate}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
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
                .getAllSites(spoAdminUrl, filter, sites.NextStartIndexFromSharePoint, webTemplate, formDigest, deleted, logger)
                .then(_ => resolve(), err => reject(err));
            }
            else {
              resolve();
            }
          }
        }, err => reject(err));
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--type [type]',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '-f, --filter [filter]'
      },
      {
        option: '--deleted'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.type) {
      if (args.options.type !== 'TeamSite' &&
        args.options.type !== 'CommunicationSite') {
        return `${args.options.type} is not a valid modern site type. Allowed types are TeamSite and CommunicationSite`;
      }
    }

    return true;
  }
}

module.exports = new SpoSiteListCommand();