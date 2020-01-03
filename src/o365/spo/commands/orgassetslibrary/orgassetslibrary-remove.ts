import {
  ContextInfo, ClientSvcResponse, ClientSvcResponseContents
} from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandValidate,
  CommandOption,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    const removeLibrary: () => void = (): void => {
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
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="RemoveFromOrgAssets" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">${args.options.libraryUrl}</Parameter><Parameter Type="Guid">{00000000-0000-0000-0000-000000000000}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            if (args.options.output === 'json') {
              cmd.log(json[json.length - 1]);
            }

            if (this.verbose) {
              cmd.log(vorpal.chalk.green('DONE'));
            }
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeLibrary();
    }
    else {
      cmd.prompt({
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

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.libraryUrl) {
        return 'Required parameter libraryUrl missing';
      }

      return true
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--libraryUrl <libraryUrl>',
        description: 'The server relative URL of the library to be removed as a central location for organization assets'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the organization asset library'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.ORGASSETSLIBRARY_REMOVE).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.

  Examples:

    Removes organization assets library without confirmation
      ${commands.ORGASSETSLIBRARY_REMOVE} --libraryUrl "/sites/branding/assets" --confirm
  `);
  }
}

module.exports = new SpoOrgAssetsLibraryRemoveCommand();
