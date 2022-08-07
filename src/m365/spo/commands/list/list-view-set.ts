import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, formatting, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  id?: string;
  title?: string;
}

class SpoListViewSetCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LIST_VIEW_SET;
  }

  public get description(): string {
    return 'Updates existing list view';
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
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
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
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--id [id]'
      },
      {
        option: '--title [title]'
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

        if (!args.options.listId && !args.options.listTitle) {
          return `Specify listId or listTitle`;
        }

        if (args.options.listId && args.options.listTitle) {
          return `Specify listId or listTitle but not both`;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        if (!args.options.id && !args.options.title) {
          return `Specify id or title`;
        }

        if (args.options.id && args.options.title) {
          return `Specify id or title but not both`;
        }

        if (args.options.id &&
          !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} in option id is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const baseRestUrl: string = `${args.options.webUrl}/_api/web/lists`;
    const listRestUrl: string = args.options.listId ?
      `(guid'${formatting.encodeQueryParameter(args.options.listId)}')`
      : `/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
    const viewRestUrl: string = `/views/${(args.options.id ? `getById('${formatting.encodeQueryParameter(args.options.id)}')` : `getByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')`)}`;

    spo
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<void> => {
        const requestOptions: any = {
          url: `${baseRestUrl}${listRestUrl}${viewRestUrl}`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: this.getPayload(args.options)
        };

        return request.patch(requestOptions);
      })
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private getPayload(options: any): any {
    const payload: any = {};
    const excludeOptions: string[] = [
      'webUrl',
      'listId',
      'listTitle',
      'id',
      'title',
      'debug',
      'verbose',
      'output'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        payload[key] = options[key];
      }
    });

    return payload;
  }
}

module.exports = new SpoListViewSetCommand();