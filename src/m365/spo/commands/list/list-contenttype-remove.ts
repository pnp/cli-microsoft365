import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  contentTypeId: string;
  confirm?: boolean;
}

class SpoListContentTypeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_CONTENTTYPE_REMOVE;
  }

  public get description(): string {
    return 'Removes content type from list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        confirm: (!(!args.options.confirm)).toString()
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
        option: '-c, --contentTypeId <contentTypeId>'
      },
      {
        option: '--confirm'
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
    
        if (args.options.listId) {
          if (!validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
          }
        }
    
        if (args.options.listId && args.options.listTitle) {
          return 'Specify listId or listTitle, but not both';
        }
    
        if (!args.options.listId && !args.options.listTitle) {
          return 'Specify listId or listTitle, one is required';
        }
    
        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('contentTypeId', 'c');
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeContentTypeFromList: () => void = (): void => {
      if (this.verbose) {
        const list: string = (args.options.listId ? args.options.listId : args.options.listTitle) as string;
        logger.logToStderr(`Removing content type ${args.options.contentTypeId} from list ${list} in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.listId) {
        requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/ContentTypes('${encodeURIComponent(args.options.contentTypeId)}')`;
      }
      else {
        requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/ContentTypes('${encodeURIComponent(args.options.contentTypeId)}')`;
      }

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .post(requestOptions)
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeContentTypeFromList();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the content type ${args.options.contentTypeId} from the list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeContentTypeFromList();
        }
      });
    }
  }
}

module.exports = new SpoListContentTypeRemoveCommand();