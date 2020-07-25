import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { HubSiteProperties } from './HubSiteProperties';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  title?: string;
  description?: string;
  logoUrl?: string;
}

class SpoHubSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_SET;
  }

  public get description(): string {
    return 'Updates properties of the specified hub site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.title = typeof args.options.title === 'string';
    telemetryProps.description = typeof args.options.description === 'string'
    telemetryProps.logoUrl = typeof args.options.logoUrl === 'string'
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Updating hub site ${args.options.id}...`);
        }

        const title: string = typeof args.options.title === 'string' ? `<SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">${Utils.escapeXml(args.options.title)}</Parameter></SetProperty>` : '';
        const description: string = typeof args.options.description === 'string' ? `<SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">${Utils.escapeXml(args.options.description)}</Parameter></SetProperty>` : '';
        const logoUrl: string = typeof args.options.logoUrl === 'string' ? `<SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">${Utils.escapeXml(args.options.logoUrl)}</Parameter></SetProperty>` : '';

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query>${title}${logoUrl}${description}<Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">${Utils.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
          const hubSite: HubSiteProperties = json.pop();
          delete hubSite._ObjectType_;

          hubSite.ID = hubSite.ID.replace('/Guid(','').replace(')/','');
          hubSite.SiteId = hubSite.SiteId.replace('/Guid(','').replace(')/','');

          cmd.log(hubSite);

          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the hub site to update'
      },
      {
        option: '-t, --title [title]',
        description: 'The new title for the hub site'
      },
      {
        option: '-d, --description [description]',
        description: 'The new description for the hub site'
      },
      {
        option: '-l, --logoUrl [logoUrl]',
        description: 'The URL of the new logo for the hub site'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (!args.options.title &&
        !args.options.description &&
        !args.options.logoUrl) {
        return 'Specify title, description or logoUrl to update';
      }

      return true;
    };
  }

  public types(): CommandTypes {
    // required to support passing empty strings as valid values
    return {
      string: ['t', 'title', 'd', 'description', 'l', 'logoUrl']
    }
  }
}

module.exports = new SpoHubSiteSetCommand();