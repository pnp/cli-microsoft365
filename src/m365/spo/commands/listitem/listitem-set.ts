import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { basic } from '../../../../utils/basic.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  id: string;
  contentType?: string;
  systemUpdate?: boolean;
}

class SpoListItemSetCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_SET;
  }

  public get description(): string {
    return 'Updates a list item in the specified list';
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
        systemUpdate: typeof args.options.systemUpdate !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id <id>'
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
        option: '-s, --systemUpdate'
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
      'id',
      'contentType'
    );
    this.types.boolean.push('systemUpdate');
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let contentTypeName: string = '';

    try {
      let requestUrl = `${args.options.webUrl}/_api/web`;

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

      if (args.options.contentType) {
        if (this.verbose) {
          await logger.logToStderr(`Getting content types for list...`);
        }

        const requestOptions: any = {
          url: `${requestUrl}/contenttypes?$select=Name,Id`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const contentTypes: any = await request.get(requestOptions);

        if (this.debug) {
          await logger.logToStderr('content type lookup response...');
          await logger.logToStderr(contentTypes);
        }

        const foundContentType: { Name: string; }[] = await basic.asyncFilter(contentTypes.value, async (ct: any) => {
          const contentTypeMatch: boolean = ct.Id.StringValue === args.options.contentType || ct.Name === args.options.contentType;

          if (this.debug) {
            await logger.logToStderr(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
          }

          return contentTypeMatch;
        });

        if (this.debug) {
          await logger.logToStderr('content type filter output...');
          await logger.logToStderr(foundContentType);
        }

        if (foundContentType.length > 0) {
          contentTypeName = foundContentType[0].Name;
        }

        // After checking for content types, throw an error if the name is blank
        if (!contentTypeName || contentTypeName === '') {
          throw `Specified content type '${args.options.contentType}' doesn't exist on the target list`;
        }

        if (this.debug) {
          await logger.logToStderr(`using content type name: ${contentTypeName}`);
        }
      }

      const options = this.mapRequestBody(args.options);

      const item = args.options.systemUpdate ?
        await spo.systemUpdateListItem(args.options.webUrl, requestUrl, args.options.id, options, contentTypeName, logger, this.verbose)
        : await spo.updateListItem(requestUrl, args.options.id, options, contentTypeName);
      delete item.ID;
      logger.log(item);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const filteredData: { [key: string]: any } = {};
    const excludeOptions: string[] = [
      'listTitle',
      'listId',
      'listUrl',
      'webUrl',
      'id',
      'contentType',
      'systemUpdate',
      'debug',
      'verbose',
      'output',
      's',
      'i',
      'o',
      'u',
      't',
      '_'
    ];

    for (const key of Object.keys(options)) {
      if (!excludeOptions.includes(key)) {
        filteredData[key] = options[key];
      }
    }

    return filteredData;
  }
}

export default new SpoListItemSetCommand();