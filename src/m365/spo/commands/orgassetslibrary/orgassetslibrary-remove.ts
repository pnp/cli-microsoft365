import { Cli, Logger } from '../../../../cli';
import {
  CommandError, CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  libraryUrl: string;
  confirm?: boolean;
}

class SpoOrgAssetsLibraryRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.ORGASSETSLIBRARY_REMOVE;
  }

  public get description(): string {
    return 'Removes a library that was designated as a central location for organization assets across the tenant.';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    const removeLibrary: () => void = (): void => {
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
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="RemoveFromOrgAssets" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">${args.options.libraryUrl}</Parameter><Parameter Type="Guid">{00000000-0000-0000-0000-000000000000}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            logger.log(json[json.length - 1]);
          }
          
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeLibrary();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the library ${args.options.libraryUrl} as a central location for organization assets?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeLibrary();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--libraryUrl <libraryUrl>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoOrgAssetsLibraryRemoveCommand();
