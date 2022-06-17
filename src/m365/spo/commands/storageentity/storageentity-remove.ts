import { Cli, Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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
    return commands.STORAGEENTITY_REMOVE;
  }

  public get description(): string {
    return 'Removes tenant property stored on the specified SharePoint Online app catalog';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --appCatalogUrl <appCatalogUrl>'
      },
      {
        option: '-k, --key <key>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.appCatalogUrl)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const removeTenantProperty = (): void => {
      let spoAdminUrl: string = '';

      if (this.verbose) {
        logger.logToStderr(`Removing tenant property ${args.options.key} from ${args.options.appCatalogUrl}...`);
      }

      spo
        .getSpoAdminUrl(logger, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return spo.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.key)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`
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
    };

    if (args.options.confirm) {
      removeTenantProperty();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to delete the ${args.options.key} tenant property?`
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
}

module.exports = new SpoStorageEntityRemoveCommand();