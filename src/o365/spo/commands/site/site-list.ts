import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { SPOSitePropertiesEnumerable } from './SPOSitePropertiesEnumerable';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.
   
  Remarks:

    Using the ${chalk.blue('-f, --filter')} option you can specify which sites you want to retrieve.
    For example, to get sites with ${chalk.grey('project')} in their URL, use ${chalk.grey("Url -like 'project'")}
    as the filter.

    When using the text output type (default), the command lists only the values
    of the ${chalk.grey('Title')}, and ${chalk.grey('Url')} properties of the site. When setting the output type to JSON,
    all available properties are included in the command output.
  
  Examples:
  
    List all modern team sites in the tenant you're logged in to
      ${commands.SITE_LIST}

    List all modern team sites in the tenant you're logged in to
      ${commands.SITE_LIST} --type TeamSite

    List all modern communication sites in the tenant you're logged in to
      ${commands.SITE_LIST} --type CommunicationSite

    List all modern team sites that contain 'project' in the URL
      ${commands.SITE_LIST} --type TeamSite --filter "Url -like 'project'"

    List all deleted sites in the tenant you're logged in to
      ${commands.SITE_LIST} --deleted
`);
  }
}

module.exports = new SpoSiteListCommand();