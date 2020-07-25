import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { SPOSitePropertiesEnumerable } from './SPOSitePropertiesEnumerable';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  filter?: string;
  deleted?: boolean;
}

class SpoSiteListCommand extends SpoCommand {
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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const siteType: string = args.options.type || 'TeamSite';
    const webTemplate: string = siteType === 'TeamSite' ? 'GROUP#0' : 'SITEPAGEPUBLISHING#0';
    let startIndex: string = '0';
    let spoAdminUrl: string;

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Retrieving list of site collections...`);
        }

        let requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">${Utils.escapeXml(args.options.filter || '')}</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">${webTemplate}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
        if (args.options.deleted) {
          requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties><Property Name="NextStartIndexFromSharePoint" ScalarProperty="true" /></Properties></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="GetDeletedSitePropertiesFromSharePoint"><Parameters><Parameter Type="Null" /></Parameters></Method></ObjectPaths></Request>`;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: requestBody
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        else {
          const sites: SPOSitePropertiesEnumerable = json[json.length - 1];
          if (args.options.output === 'json') {
            cmd.log(sites._Child_Items_);
          }
          else {
            cmd.log(sites._Child_Items_.map(s => {
              return {
                Title: s.Title,
                Url: s.Url
              };
            }).sort((a, b) => {
              const urlA = a.Url.toUpperCase();
              const urlB = b.Url.toUpperCase();
              if (urlA < urlB) {
                return -1;
              }
              if (urlA > urlB) {
                return 1;
              }

              return 0;
            }));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--type [type]',
        description: 'type of modern sites to list. Allowed values TeamSite|CommunicationSite, default TeamSite',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '-f, --filter [filter]',
        description: 'filter to apply when retrieving sites'
      },
      {
        option: '--deleted',
        description: 'use this switch to only return deleted sites'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.type) {
        if (args.options.type !== 'TeamSite' &&
          args.options.type !== 'CommunicationSite') {
          return `${args.options.type} is not a valid modern site type. Allowed types are TeamSite and CommunicationSite`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoSiteListCommand();