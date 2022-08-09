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
  url: string;
  confirm?: boolean;
}

class SpoOrgNewsSiteRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.ORGNEWSSITE_REMOVE;
  }

  public get description(): string {
    return 'Removes a site from the list of organizational news sites';
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
        option: '-u, --url <url>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const removeOrgNewsSite = (): void => {
      let spoAdminUrl: string = '';

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
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="64" ObjectPathId="63" /><Method Name="RemoveOrgNewsSite" Id="65" ObjectPathId="63"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="63" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            cb();
          }
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeOrgNewsSite();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove ${args.options.url} from the list of organizational news sites?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeOrgNewsSite();
        }
      });
    }
  }
}

module.exports = new SpoOrgNewsSiteRemoveCommand();