import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  confirm?: boolean;
}

class SpoHubSiteDisconnectCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_DISCONNECT;
  }

  public get description(): string {
    return 'Disconnects the specifies site collection from its hub site';
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
        confirm: args.options.confirm || false
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const disconnectHubSite: () => void = (): void => {
      spo
        .getRequestDigest(args.options.url)
        .then((res: ContextInfo): Promise<void> => {
          if (this.verbose) {
            logger.logToStderr(`Disconnecting site collection ${args.options.url} from its hubsite...`);
          }

          const requestOptions: any = {
            url: `${args.options.url}/_api/site/JoinHubSite('00000000-0000-0000-0000-000000000000')`,
            headers: {
              'X-RequestDigest': res.FormDigestValue,
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          return request.post(requestOptions);
        })
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      disconnectHubSite();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to disconnect the site collection ${args.options.url} from its hub site?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          disconnectHubSite();
        }
      });
    }
  }
}

module.exports = new SpoHubSiteDisconnectCommand();