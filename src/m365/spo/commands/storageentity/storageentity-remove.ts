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
  appCatalogUrl: string;
  key: string;
  confirm?: boolean;
}

class SpoStorageEntityRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.STORAGEENTITY_REMOVE}`;
  }

  public get description(): string {
    return 'Removes tenant property stored on the specified SharePoint Online app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const removeTenantProperty = (): void => {
      let spoAdminUrl: string = '';

      if (this.verbose) {
        cmd.log(`Removing tenant property ${args.options.key} from ${args.options.appCatalogUrl}...`);
      }

      this
        .getSpoAdminUrl(cmd, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return this.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.key)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`
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
    }

    if (args.options.confirm) {
      removeTenantProperty();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to delete the ${args.options.key} tenant property?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeTenantProperty();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --appCatalogUrl <appCatalogUrl>',
        description: 'URL of the app catalog site'
      },
      {
        option: '-k, --key <key>',
        description: 'Name of the tenant property to retrieve'
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
      const result: boolean | string = SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      if (result === false) {
        return 'Missing required option appCatalogUrl';
      }
      else {
        return result;
      }
    };
  }
}

module.exports = new SpoStorageEntityRemoveCommand();