import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  secondary?: boolean;
  confirm?: boolean;
}

class SpoSiteRecycleBinItemClearCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Permanently removes all items in a site recycle bin';
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
        secondary: !!args.options.secondary,
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
        option: '--secondary'
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

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.clearRecycleBin(args, logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear the recycle bin of site ${args.options.siteUrl}?`
      });

      if (result.continue) {
        await this.clearRecycleBin(args, logger);
      }
    }
  }

  private async clearRecycleBin(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Permanently removing all items in recycle bin of site ${args.options.siteUrl}...`);
      }

      const requestOptions: CliRequestOptions = {
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.secondary) {
        if (this.verbose) {
          logger.logToStderr('Removing all items from the second-stage recycle bin');
        }
        requestOptions.url = `${args.options.siteUrl}/_api/site/RecycleBin/DeleteAllSecondStageItems`;
      }
      else {
        if (this.verbose) {
          logger.logToStderr('Removing all items from the first-stage recycle bin');
        }
        requestOptions.url = `${args.options.siteUrl}/_api/web/RecycleBin/DeleteAll`;
      }

      const result = await request.post<any>(requestOptions);
      if (result['odata.null'] !== true) {
        throw result;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteRecycleBinItemClearCommand();