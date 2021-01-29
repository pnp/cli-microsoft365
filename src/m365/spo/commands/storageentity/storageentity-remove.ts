import { Cli, Logger } from '../../../../cli';
import {
  CommandError, CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../../spo';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const removeTenantProperty = (): void => {
      let spoAdminUrl: string = '';

      if (this.verbose) {
        logger.logToStderr(`Removing tenant property ${args.options.key} from ${args.options.appCatalogUrl}...`);
      }

      this
        .getSpoAdminUrl(logger, this.debug)
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
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.key)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`
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
            cb();
          }
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    }

    if (args.options.confirm) {
      removeTenantProperty();
    }
    else {
      Cli.prompt({
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
        option: '-u, --appCatalogUrl <appCatalogUrl>'
      },
      {
        option: '-k, --key <key>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const result: boolean | string = SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
    if (result === false) {
      return 'Missing required option appCatalogUrl';
    }
    else {
      return result;
    }
  }
}

module.exports = new SpoStorageEntityRemoveCommand();