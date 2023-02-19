import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
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
  contentTypeId: string;
}

interface ContentTypes {
  value: {
    Id: StringValue;
  }[];
}

interface StringValue {
  StringValue: string;
}

class SpoListContentTypeDefaultSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_CONTENTTYPE_DEFAULT_SET;
  }

  public get description(): string {
    return 'Sets the default content type for a list';
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
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-c, --contentTypeId <contentTypeId>'
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

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('contentTypeId', 'c');
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    let baseUrl: string = `${args.options.webUrl}/_api/web/`;
    if (args.options.listId) {
      baseUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      baseUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      baseUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    if (this.verbose) {
      logger.logToStderr('Retrieving content type order...');
    }

    try {
      const contentTypeOrder: StringValue[] = await this.getContentTypeOrder(baseUrl, logger);

      // see if the specified content type is among the registered content types
      // if it is, it means it's visible
      const contentTypeIndex: number = contentTypeOrder.findIndex(ct => ct.StringValue.toUpperCase() === args.options.contentTypeId.toUpperCase());

      if (contentTypeIndex > -1) {
        if (this.debug) {
          logger.logToStderr(`Content type ${args.options.contentTypeId} is visible in the list`);
        }
        // content type is in the list and is visible in the menu

        if (contentTypeIndex === 0) {
          if (this.verbose) {
            logger.logToStderr(`Content type ${args.options.contentTypeId} is already set as default`);
          }
        }
        else {
          if (this.verbose) {
            logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
          }

          // remove content type from the order array so that we can put it at
          // the beginning to make it default content type          
          contentTypeOrder.splice(contentTypeIndex, 1);
          contentTypeOrder.unshift({
            StringValue: args.options.contentTypeId
          });

          await this.updateContentTypeOrder(baseUrl, contentTypeOrder);
        }
      }
      else {
        if (this.debug) {
          logger.logToStderr(`Content type ${args.options.contentTypeId} is not visible in the list`);
        }

        if (this.verbose) {
          logger.logToStderr('Retrieving list content types...');
        }

        const contentTypes: string[] = await this.getListContentTypes(baseUrl);
        if (!contentTypes.find(ct => ct.toUpperCase() === args.options.contentTypeId.toUpperCase())) {
          throw `Content type ${args.options.contentTypeId} missing in the list. Add the content type to the list first and try again.`;
        }

        if (this.verbose) {
          logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
        }

        contentTypeOrder.unshift({
          StringValue: args.options.contentTypeId
        });

        await this.updateContentTypeOrder(baseUrl, contentTypeOrder);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContentTypeOrder(baseUrl: string, logger: Logger): Promise<StringValue[]> {
    const requestOptions: CliRequestOptions = {
      url: `${baseUrl}/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ ContentTypeOrder: StringValue[]; UniqueContentTypeOrder: StringValue[] | null }>(requestOptions);
    let uniqueContentTypeOrder = response.ContentTypeOrder;
    if (response.UniqueContentTypeOrder !== null) {
      if (this.debug) {
        logger.logToStderr('Using unique content type order');
      }
      uniqueContentTypeOrder = response.UniqueContentTypeOrder as StringValue[];
    }
    else {
      if (this.debug) {
        logger.logToStderr('Unique content type order not defined. Using content type order');
      }
    }
    return uniqueContentTypeOrder;
  }

  private async updateContentTypeOrder(baseUrl: string, contentTypeOrder: StringValue[]): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${baseUrl}/RootFolder`,
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'x-http-method': 'MERGE'
      },
      data: {
        UniqueContentTypeOrder: contentTypeOrder
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async getListContentTypes(baseUrl: string): Promise<string[]> {
    const requestOptions: CliRequestOptions = {
      url: `${baseUrl}/ContentTypes?$select=Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    const response = await request.get<ContentTypes>(requestOptions);
    return response.value.map(ct => ct.Id.StringValue);
  }
}

module.exports = new SpoListContentTypeDefaultSetCommand();