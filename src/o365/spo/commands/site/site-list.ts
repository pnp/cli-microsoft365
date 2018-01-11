import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import SpoCommand from '../../SpoCommand';
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
}

class SiteListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_LIST;
  }

  public get description(): string {
    return 'Lists modern sites of the given type';
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.siteType = args.options.type || 'TeamSite';
    telemetryProps.filter = (!(!args.options.filter)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const siteType: string = args.options.type || 'TeamSite';
    const webTemplate: string = siteType === 'TeamSite' ? 'GROUP#0' : 'SITEPAGEPUBLISHING#0';
    let startIndex: string = '0';

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<ContextInfo> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest for tenant admin at ${auth.site.url}...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Retrieving list of site collections...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">${Utils.escapeXml(args.options.filter || '')}</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">${webTemplate}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(JSON.stringify(requestOptions));
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cmd.log(new CommandError(response.ErrorInfo.ErrorMessage));
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
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
      using the ${chalk.blue(commands.CONNECT)} command.
   
  Remarks:

    To list modern sites, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
    If you are connected to a different site and will try to list the available sites,
    you will get an error.

    Using the ${chalk.blue('-f, --filter')} option you can specify which sites you want to retrieve.
    For example, to get sites with ${chalk.grey('project')} in their URL, use ${chalk.grey("Url -like 'project'")}
    as the filter.

    When using the text output type (default), the command lists only the values
    of the ${chalk.grey('Title')}, and ${chalk.grey('Url')} properties of the site. When setting the output type to JSON,
    all available properties are included in the command output.
  
  Examples:
  
    List all modern team sites in the currently connected tenant
      ${chalk.grey(config.delimiter)} ${commands.SITE_LIST}

    List all modern team sites in the currently connected tenant
      ${chalk.grey(config.delimiter)} ${commands.SITE_LIST} --type TeamSite

    List all modern communication sites in the currently connected tenant
      ${chalk.grey(config.delimiter)} ${commands.SITE_LIST} --type CommunicationSite

    List all modern team sites that contain 'project' in the URL
      ${chalk.grey(config.delimiter)} ${commands.SITE_LIST} --type TeamSite --filter "Url -like 'project'"
`);
  }
}

module.exports = new SiteListCommand();