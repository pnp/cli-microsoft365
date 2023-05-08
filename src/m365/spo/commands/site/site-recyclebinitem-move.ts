// import { v4 } from 'uuid';
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
  ids?: string;
  all?: boolean;
  confirm?: boolean;
}

class SpoSiteRecycleBinItemMoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_MOVE;
  }

  public get description(): string {
    return 'Moves items from the first-stage recycle bin to the second-stage recycle bin';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        ids: typeof args.options.ids !== 'undefined',
        all: !!args.options.all,
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
        option: '-i, --ids [ids]'
      },
      {
        option: '--all'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['ids', 'all'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.ids && !validation.isValidGuidArray(args.options.ids.split(','))) {
          return 'The option ids contains one or more invalid GUIDs';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Moving items from the first-stage recycle bin to the second-stage recycle bin at ${args.options.siteUrl}...`);
    }

    if (args.options.confirm) {
      await this.moveRecycleBinItem(args, logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: 'Are you sure you want to move these items to the second-stage recycle bin?'
      });

      if (result.continue) {
        await this.moveRecycleBinItem(args, logger);
      }
    }
  }

  private async moveRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      if (args.options.all) {
        if (this.verbose) {
          logger.logToStderr('Moving all items to the second-stage recycle bin');
        }
        const requestOptions: CliRequestOptions = {
          url: `${args.options.siteUrl}/_api/web/recycleBin/MoveAllToSecondStage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        await request.post<{ 'odata.null': boolean }>(requestOptions);
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Moving ${args.options.ids} to the second-stage recycle bin`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${args.options.siteUrl}/_api/web/recycleBin/MoveAllToSecondStage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            'ids': args.options.ids!.split(',')
          }
        };
        await request.post<{ 'odata.null': boolean }>(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteRecycleBinItemMoveCommand();