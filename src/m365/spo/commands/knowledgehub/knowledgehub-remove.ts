import { Cli, Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
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
        option: '--confirm'
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    const removeKnowledgehub = (): void => {
      spo
        .getSpoAdminUrl(logger, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return spo.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            logger.logToStderr(`Removing Knowledge Hub Site settings from your tenant`);
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
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
            logger.log(json[json.length - 1]);
            cb();
          }
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      if (this.debug) {
        logger.logToStderr('Confirmation bypassed by entering confirm option. Removing Knowledge Hub Site setting...');
      }
      removeKnowledgehub();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove Knowledge Hub Site from your tenant?`
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
}

module.exports = new SpoKnowledgehubRemoveCommand();