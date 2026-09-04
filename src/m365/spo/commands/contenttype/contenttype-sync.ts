import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  name?: string;
  listTitle?: string;
  listId?: string;
  listUrl?: string;
}

class SpoContentTypeSyncCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_SYNC;
  }

  public get description(): string {
    return 'Adds a published content type from the content type hub to a site or syncs its latest changes';
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
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
          return `${args.options.listId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'id', 'name', 'listTitle', 'listId', 'listUrl');
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'name']
      },
      {
        options: ['listId', 'listTitle', 'listUrl'],
        runsWhen: (args) => args.options.listId || args.options.listTitle || args.options.listUrl
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { listId, listTitle, listUrl, webUrl } = args.options;
    const url: URL = new URL(webUrl);
    const baseUrl = `${url.origin}/_api/v2.0/sites/`;

    try {
      const [siteId, webId] = await Promise.all([
        spo.getSiteIdBySPApi(webUrl, logger, this.verbose),
        spo.getWebId(webUrl, logger, this.verbose)
      ]);
      const siteUrl = `${url.host},${siteId},${webId}`;
      const targetListId = listId || (listTitle || listUrl ? await spo.getListId(webUrl, listTitle, listUrl, logger, this.verbose) : undefined);
      const listPath = targetListId ? `/lists/${targetListId}` : '';
      const contentTypeId = await this.getContentTypeId(webUrl, args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding or syncing the content type...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${baseUrl}${siteUrl}${listPath}/contenttypes/addCopyFromContentTypeHub`,
        headers: {
          'accept': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false'
        },
        responseType: 'json',
        data: {
          contentTypeId: contentTypeId
        }
      };

      const response = await request.post(requestOptions);

      // The endpoint only returns a response if the content type has been added for the first time
      // When syncing, the response will be an empty string, which should not be logged.
      if (typeof response === 'object') {
        await logger.log(response);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContentTypeId(webUrl: string, options: Options, logger: Logger): Promise<string> {
    if (options.id) {
      return options.id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving content type Id by name...`);
    }

    const contentTypes: { StringId: string }[] = await odata.getAllItems(`${webUrl}/_api/web/AvailableContentTypes?$filter=Name eq '${formatting.encodeQueryParameter(options.name!)}'&$select=StringId,Name`);

    if (contentTypes.length === 0) {
      throw `Content type with name ${options.name} not found.`;
    }

    return contentTypes[0].StringId;
  }
}

export default new SpoContentTypeSyncCommand();