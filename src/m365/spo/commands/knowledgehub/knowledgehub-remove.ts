import commands from '../../commands';
import config from '../../../../config';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import {
  CommandOption,
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { CommandInstance } from '../../../../cli';
import * as chalk from 'chalk';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class SpoKnowledgehubRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.KNOWLEDGEHUB_REMOVE;
  }

  public get description(): string {
    return 'Removes the Knowledge Hub Site setting for your tenant';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    const removeKnowledgehub = (): void => {
      this
        .getSpoAdminUrl(cmd, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return this.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            cmd.log(`Removing Knowledge Hub Site settings from your tenant`);
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
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
            cmd.log(json[json.length - 1]);

            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
            }
            cb();
          }
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      if (this.debug) {
        cmd.log('Confirmation bypassed by entering confirm option. Removing Knowledge Hub Site setting...');
      }
      removeKnowledgehub();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove Knowledge Hub Site from your tenant?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeKnowledgehub();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removal of Knowledge Hub Site for your tenant'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoKnowledgehubRemoveCommand();