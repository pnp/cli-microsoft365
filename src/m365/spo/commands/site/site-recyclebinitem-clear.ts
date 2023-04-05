import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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

class SpoSiteRecycleBinItemListCommand extends SpoCommand {
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
    if (this.verbose) {
      logger.logToStderr(`Permanently removes all items in a site recycle bin at ${args.options.siteUrl}...`);
    }

    if (args.options.confirm) {
      await this.clearRecycleBin(args, logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear the recycle bin?`
      });

      if (result.continue) {
        await this.clearRecycleBin(args, logger);
      }
    }
  }
  private async clearRecycleBin(args: CommandArgs, logger: Logger): Promise<void> {
    const requestOptions: any = {
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (args.options.secondary) {
      if (this.verbose) {
        logger.logToStderr(`Removing items from the second-stage recycle bin`);
      }
      requestOptions.url = `${args.options.siteUrl}/_api/site/RecycleBin/DeleteAllSecondStageItems`;
    }
    else {
      if (this.verbose) {
        logger.logToStderr(`Removing items from the first-stage recycle bin`);
      }
      requestOptions.url = `${args.options.siteUrl}/_api/web/RecycleBin/DeleteAll`;
    }

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteRecycleBinItemListCommand();