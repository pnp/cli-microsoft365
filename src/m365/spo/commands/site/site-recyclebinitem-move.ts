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
    this.#initTypes();
    this.#initOptionSets();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        ids: typeof args.options.siteidsrl !== 'undefined',
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

  #initTypes(): void {
    this.types.boolean.push('all');
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

        if (args.options.ids) {
          const ids: string[] = args.options.ids.split(',').map(i => i.trim());
          for (let i = 0; i < ids.length; i++) {
            if (!validation.isValidGuid(ids[i] as string)) {
              return `${ids[i]} is not a valid GUID`;
            }
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all items from recycle bin at ${args.options.siteUrl}...`);
    }

    if (args.options.confirm) {
      await this.moveRecycleBinItem(args, logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to move the items to the second-stage recycle bin?`
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
          logger.logToStderr(`Moving all items to the second-stage recycle bin`);
        }
        const requestOptions: any = {
          url: `${args.options.siteUrl}/_api/web/recycleBin/MoveAllToSecondStage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        await request.post<{ value: any[] }>(requestOptions);
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Moving ${args.options.ids} to the second-stage recycle bin`);
        }
        const requestOptions: any = {
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const splittedIds = args.options.ids!.split(',');
        for (const id of splittedIds) {
          requestOptions.url = `${args.options.siteUrl}/_api/web/recycleBin('${id.trim()}')/MoveToSecondStage`;
          await request.post<{ value: any[] }>(requestOptions);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteRecycleBinItemMoveCommand();