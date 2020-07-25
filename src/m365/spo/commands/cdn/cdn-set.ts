import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
  enabled: string;
  noDefaultOrigins?: boolean;
}

class SpoCdnSetCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_SET;
  }

  public get description(): string {
    return 'Enable or disable the specified Microsoft 365 CDN';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.cdnType = args.options.type || 'Public';
    telemetryProps.enabled = args.options.enabled === 'true';
    telemetryProps.noDefaultOrigins = (!(!args.options.noDefaultOrigins)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const cdnTypeString: string = args.options.type || 'Public';
    const enabled: boolean = args.options.enabled === 'true';
    let cdnType: number = 0;
    let spoAdminUrl: string = '';

    switch (cdnTypeString) {
      case "Private": {
        cdnType = 1;
        break;
      }
      case "Both": {
        cdnType = 2;
        break;
      }
      default: {
        cdnType = 0;
        break;
      }
    }

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        let requestBody = '';

        if (cdnType === 2) {
          if (args.options.noDefaultOrigins) {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} Public and Private CDNs without default origins. Please wait, this might take a moment...`);
            }

            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><Method Name="SetTenantCdnEnabled" Id="13" ObjectPathId="11"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="14" ObjectPathId="11"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
          }
          else {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} Public and Private CDNs with default origins. Please wait, this might take a moment...`);
            }

            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method>${(enabled ? '<Method Name="CreateTenantCdnDefaultOrigins" Id="99" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="100" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method>' : '')}</Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
          }
        }
        else {
          if (args.options.noDefaultOrigins) {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} ${(cdnType === 1 ? 'Private' : 'Public')} CDN without default origins. Please wait, this might take a moment...`);
            }

            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
          }
          else {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
            }

            if (enabled) {
              requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">${(cdnType === 1 ? 1 : 0)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
            }
            else {
              requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
            }
          }
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
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.type) {
        if (args.options.type !== 'Public' && args.options.type !== 'Both' &&
          args.options.type !== 'Private') {
          return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private|Both`;
        }
      }

      if (args.options.enabled !== 'true' &&
        args.options.enabled !== 'false') {
        return `${args.options.enabled} is not a valid boolean value. Allowed values are true|false`;
      }

      return true;
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --enabled <enabled>',
        description: 'Set to true to enable CDN or to false to disable it. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '-t, --type [type]',
        description: 'Type of CDN to manage. Public|Private|Both. Default Public',
        autocomplete: ['Public', 'Private', 'Both']
      },
      {
        option: '--noDefaultOrigins',
        description: 'Disable creation of default origins'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoCdnSetCommand();