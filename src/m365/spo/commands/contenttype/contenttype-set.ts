import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import request from '../../../../request';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  name?: string;
  listTitle?: string;
  listId: string;
  listUrl: string;
}

class SpoContentTypeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_SET;
  }

  public get description(): string {
    return 'Updates existing content type';
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `'${args.options.listId}' is not a valid GUID.`;
        }

        if ((args.options.listId && (args.options.listTitle || args.options.listUrl)) || (args.options.listTitle && args.options.listUrl)) {
          return `Specify either listTitle, listId or listUrl.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'i');
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'name'] }
    );
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: AxiosRequestConfig = {
      url: `${args.options.webUrl}/_api/Web`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: this.getRequestPayload(args.options)
    };

    if (args.options.listId) {
      requestOptions.url += `/Lists/GetById('${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestOptions.url += `/Lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      requestOptions.url += `/GetList('${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl))}')`;
    }
    requestOptions.url += '/ContentTypes';

    try {
      const contentTypeId = await this.getContentTypeId(args.options);
      requestOptions.url += `/GetById('${formatting.encodeQueryParameter(contentTypeId)}')`;

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getContentTypeId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${options.webUrl}/_api/Web`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (options.listId) {
      requestOptions.url += `/Lists/GetById('${formatting.encodeQueryParameter(options.listId)}')`;
    }
    else if (options.listTitle) {
      requestOptions.url += `/Lists/GetByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }
    else if (options.listUrl) {
      requestOptions.url += `/GetList('${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(options.webUrl, options.listUrl))}')`;
    }
    requestOptions.url += `/ContentTypes?$filter=Name eq '${formatting.encodeQueryParameter(options.name!)}'&$select=Id`;

    const res = await request.get<{ value: { Id: { StringValue: string } }[] }>(requestOptions);

    if (res.value.length === 0) {
      throw `The specified content type '${options.name}' does not exist`;
    }

    return res.value[0].Id.StringValue;
  }

  private getRequestPayload(options: Options): any {
    const excludeOptions: string[] = [
      'webUrl',
      'id',
      'name',
      'listTitle',
      'listId',
      'listUrl',
      'query',
      'debug',
      'verbose',
      'output'
    ];

    const payload = Object.keys(options)
      .filter(key => excludeOptions.indexOf(key) === -1)
      .reduce((object: any, key: string) => {
        object[key] = options[key];
        return object;
      }, {});

    return payload;
  }
}

module.exports = new SpoContentTypeSetCommand();