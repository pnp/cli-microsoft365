import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  force?: boolean;
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
        force: args.options.force || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const unregisterHubSite = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Unregistering site collection ${args.options.url} as a hub site...`);
        }

        const res = await spo.getRequestDigest(args.options.url);

        const requestOptions: CliRequestOptions = {
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

    if (args.options.force) {
      await unregisterHubSite();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to unregister the site collection ${args.options.url} as a hub site?`);

      if (result) {
        await unregisterHubSite();
      }
    }
  }
}

export default new SpoHubSiteUnregisterCommand();