import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CommandError } from '../../../../Command';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  ids: string;
  confirm?: boolean;
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
        confirm: !!args.options.confirm
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
        option: '--confirm'
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
          return 'The option ids contains one or more invalid GUIDs';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeRecycleBinItem(args, logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to permanently delete ${args.options.ids.split(',').length} item(s) from the site recycle bin?`
      });

      if (result.continue) {
        await this.removeRecycleBinItem(args, logger);
      }
    }
  }

  private async removeRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Permanently deleting specific items from the site recycle bin at ${args.options.siteUrl}...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/RecycleBin/DeleteByIds`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          'ids': args.options.ids.split(',')
        }
      };
      const response = await request.post<any>(requestOptions);
      logger.log(response);
    }
    catch (err: any) {
      if (err.message && err.message === 'Request failed with status code 400') {
        throw new CommandError('Failed to remove one or more IDs from the recycle bin. Please check the ids');
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteRecycleBinItemRemoveCommand();