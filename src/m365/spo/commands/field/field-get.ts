import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  id?: string;
  fieldTitle?: string;
  title?: string;
}

class SpoFieldGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_GET;
  }

  public get description(): string {
    return 'Retrieves information about the specified list- or site column';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--fieldTitle [fieldTitle]'
      },
      {
        option: '-t, --title [title]'
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

        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'title', 'fieldTitle']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.fieldTitle) {
      args.options.title = args.options.fieldTitle;

      this.warn(logger, `Option 'fieldTitle' is deprecated. Please use 'title' instead.`);
    }

    let listRestUrl: string = '';

    if (args.options.listId) {
      listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else if (args.options.listTitle) {
      listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
    }

    let fieldRestUrl: string = '';
    if (args.options.id) {
      fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else {
      fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter(args.options.title as string)}')`;
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoFieldGetCommand();