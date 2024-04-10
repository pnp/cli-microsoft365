import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  contentTypeId?: string;
  contentTypeName?: string;
  listTitle?: string;
  listId?: string;
  listUrl?: string;
  properties?: string;
}

class SpoContentTypeFieldListCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_FIELD_LIST;
  }

  public get description(): string {
    return 'Lists fields for a given site or list content type';
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
        contentTypeId: typeof args.options.contentTypeId !== 'undefined',
        contentTypeName: typeof args.options.contentTypeName !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        properties: typeof args.options.properties !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --contentTypeId [contentTypeId]'
      },
      {
        option: '-n, --contentTypeName [contentTypeName]'
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
        option: '-p, --properties [properties]'
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

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID for option 'listId'.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'contentTypeId', 'contentTypeName', 'listTitle', 'listId', 'listUrl', 'properties');
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['contentTypeId', 'contentTypeName']
      },
      {
        options: ['listId', 'listTitle', 'listUrl'],
        runsWhen: (args) => args.options.listId || args.options.listTitle || args.options.listUrl
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving fields for content type ${args.options.contentTypeId || args.options.contentTypeName} in site ${args.options.webUrl}...`);
      }

      let requestUrl: string = `${args.options.webUrl}/_api/web`;
      if (args.options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      requestUrl += '/contentTypes';

      const contentTypeId = await this.getContentTypeId(args.options.webUrl, requestUrl, logger, args.options.contentTypeId, args.options.contentTypeName);
      requestUrl += `('${formatting.encodeQueryParameter(contentTypeId)}')/fields`;

      if (args.options.properties) {
        requestUrl += `?$select=${args.options.properties}`;
      }

      const res = await odata.getAllItems(requestUrl);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContentTypeId(webUrl: string, requestUrl: string, logger: Logger, contentTypeId?: string, contentTypeName?: string): Promise<string> {
    if (contentTypeId) {
      return contentTypeId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving content type id for contentType ${contentTypeName}.`);
    }

    const res: { StringId: string }[] = await odata.getAllItems(`${requestUrl}?$filter=Name eq '${formatting.encodeQueryParameter(contentTypeName!)}'&$select=StringId`);

    if (res.length === 0) {
      throw `Content type with name ${contentTypeName} not found.`;
    }

    return res[0].StringId;
  }
}

export default new SpoContentTypeFieldListCommand();