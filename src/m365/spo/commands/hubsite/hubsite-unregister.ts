import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  confirm?: boolean;
}

class SpoHubSiteUnregisterCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_UNREGISTER;
  }

  public get description(): string {
    return 'Unregisters the specified site collection as a hub site';
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
    const unregisterHubSite: () => Promise<void> = async (): Promise<void> => {
      try {
        const res = await spo.getRequestDigest(args.options.url);

        const requestOptions: any = {
          url: `${args.options.url}/_api/site/UnregisterHubSite`,
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
      await unregisterHubSite();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to unregister the site collection ${args.options.url} as a hub site?`
      });

      if (result.continue) {
        await unregisterHubSite();
      }
    }
  }
}

module.exports = new SpoHubSiteUnregisterCommand();