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
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
  origin: string;
  confirm?: boolean;
}

class SpoCdnOriginRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_ORIGIN_REMOVE;
  }

  public get description(): string {
    return 'Removes CDN origin for the current SharePoint Online tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.cdnType = args.options.type || 'Public';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    const removeCdnOrigin = (): void => {
      this
        .getTenantId(cmd, this.debug)
        .then((_tenantId: string): Promise<string> => {
          tenantId = _tenantId;
          return this.getSpoAdminUrl(cmd, this.debug);
        })
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return this.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            cmd.log(`Removing origin ${args.options.origin} from the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.origin)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="${tenantId}" /></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
          }
          else {
            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
            }
            cb();
          }
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      if (this.debug) {
        cmd.log('Confirmation suppressed through the confirm option. Removing CDN origin...');
      }
      removeCdnOrigin();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to delete the ${args.options.origin} CDN origin?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeCdnOrigin();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --type [type]',
        description: 'Type of CDN to manage. Public|Private. Default Public',
        autocomplete: ['Public', 'Private']
      },
      {
        option: '-r, --origin <origin>',
        description: 'Origin to remove from the current CDN configuration'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removal of a tenant property'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.type) {
        if (args.options.type !== 'Public' &&
          args.options.type !== 'Private') {
          return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoCdnOriginRemoveCommand();