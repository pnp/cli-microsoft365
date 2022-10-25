import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
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
  viewId?: string;
  viewTitle?: string;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        viewId: typeof args.options.viewId !== 'undefined',
        viewTitle: typeof args.options.viewTitle !== 'undefined'
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
        option: '--listUrl [listUrl]'
      },
      {
        option: '--viewId [viewId]'
      },
      {
        option: '--viewTitle [viewTitle]'
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

        if (args.options.viewId &&
          !validation.isValidGuid(args.options.viewId)) {
          return `${args.options.viewId} in option viewId is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      ['listId', 'listTitle', 'listUrl'],
      ['viewId', 'viewTitle']
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let listRestUrl: string = '';
    if (args.options.listId) {
      listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      listRestUrl = `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const viewRestUrl: string = `/views/${(args.options.viewId ? `GetById('${formatting.encodeQueryParameter(args.options.viewId)}')` : `GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`)}`;

    try {
      const res = await spo.getRequestDigest(args.options.webUrl);
      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/${listRestUrl}${viewRestUrl}`,
        headers: {
          'X-RequestDigest': res.FormDigestValue,
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: this.getPayload(args.options)
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getPayload(options: any): any {
    const payload: any = {};
    const excludeOptions: string[] = [
      'webUrl',
      'listId',
      'listTitle',
      'listUrl',
      'viewId',
      'viewTitle',
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