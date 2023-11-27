import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  ids: string;
  force?: boolean;
}

class SpoSiteRecycleBinItemRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Permanently deletes specific items from the site recycle bin';
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
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --ids <ids>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (!validation.isValidGuidArray(args.options.ids.split(','))) {
          return 'The option ids contains one or more invalid GUIDs.';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRecycleBinItem(args, logger);
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to permanently delete ${args.options.ids.split(',').length} item(s) from the site recycle bin?` });

      if (result) {
        await this.removeRecycleBinItem(args, logger);
      }
    }
  }

  private async removeRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Permanently deleting items from the site recycle bin at site ${args.options.siteUrl}...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/RecycleBin/DeleteByIds`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          ids: args.options.ids.split(',')
        }
      };

      await request.post<any>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteRecycleBinItemRemoveCommand();