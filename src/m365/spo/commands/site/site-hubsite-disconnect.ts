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
  siteUrl: string;
  force?: boolean;
}

class SpoSiteHubSiteDisconnectCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_HUBSITE_DISCONNECT;
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
        force: args.options.force || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.disconnectHubSite(logger, args);
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to disconnect the site collection ${args.options.siteUrl} from its hub site?`);

      if (result) {
        await this.disconnectHubSite(logger, args);
      }
    }
  }

  private async disconnectHubSite(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const res = await spo.getRequestDigest(args.options.siteUrl);

      if (this.verbose) {
        await logger.logToStderr(`Disconnecting site collection ${args.options.siteUrl} from its hubsite...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/JoinHubSite('00000000-0000-0000-0000-000000000000')`,
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
  }
}

export default new SpoSiteHubSiteDisconnectCommand();