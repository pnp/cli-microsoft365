import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { ListItemAddOptions, spoListItem } from '../../../../utils/spoListItem.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  contentType?: string;
  folder?: string;
}

class SpoListItemAddCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_ADD;
  }

  public get description(): string {
    return 'Creates a list item in the specified list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        contentType: typeof args.options.contentType !== 'undefined',
        folder: typeof args.options.folder !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-c, --contentType [contentType]'
      },
      {
        option: '--folder [folder]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push(
      'webUrl',
      'listId',
      'listTitle',
      'listUrl',
      'contentType',
      'folder'
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const options: ListItemAddOptions = {
        webUrl: args.options.webUrl,
        listId: args.options.listId,
        listUrl: args.options.listUrl,
        listTitle: args.options.listTitle,
        contentType: args.options.contentType,
        folder: args.options.folder,
        fieldValues: this.mapUnknownProperties(args.options)
      };

      const item = await spoListItem.addListItem(options, logger, this.verbose, this.debug);
      await logger.log(item);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapUnknownProperties(options: Options): any {
    const fieldValues: { [key: string]: any } = {};
    const excludeOptions: string[] = [
      'listTitle',
      'listId',
      'listUrl',
      'webUrl',
      'contentType',
      'folder',
      'debug',
      'verbose',
      'output',
      '_'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        fieldValues[key] = `${(<any>options)[key]}`;
      }
    });

    return fieldValues;
  }
}

export default new SpoListItemAddCommand();