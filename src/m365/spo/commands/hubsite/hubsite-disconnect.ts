import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const disconnectHubSite: () => Promise<void> = async (): Promise<void> => {
      try {
        const res = await spo.getRequestDigest(args.options.url);

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

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await disconnectHubSite();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to disconnect the site collection ${args.options.url} from its hub site?`
      });

      if (result.continue) {
        await disconnectHubSite();
      }
    }
  }
}

module.exports = new SpoHubSiteDisconnectCommand();