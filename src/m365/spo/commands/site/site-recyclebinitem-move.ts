import { cli } from '../../../../cli/cli.js';
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
  ids?: string;
  all?: boolean;
  force?: boolean;
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
        option: '-i, --ids [ids]'
      },
      {
        option: '--all'
      },
      {
        option: '-f, --force'
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
    if (args.options.force) {
      await this.moveRecycleBinItem(args, logger);
    }
    else {
      const result = await cli.promptForConfirmation({ message: 'Are you sure you want to move these items to the second-stage recycle bin?' });

      if (result) {
        await this.moveRecycleBinItem(args, logger);
      }
    }
  }

  private async moveRecycleBinItem(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/web/recycleBin/MoveAllToSecondStage`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.ids !== undefined) {
        const ids = args.options.ids.split(',');
        if (this.verbose) {
          await logger.logToStderr(`Moving ${ids.length} items to the second-stage recycle bin.`);
        }

        requestOptions.data = {
          ids: ids
        };
      }
      else if (this.verbose) {
        await logger.logToStderr('Moving all items to the second-stage recycle bin.');
      }

      await request.post<any>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteRecycleBinItemMoveCommand();