import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstance } from './ListItemInstance';

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

interface FieldValue {
  ErrorMessage: string;
  FieldName: string;
  FieldValue: any;
  HasException: boolean;
  ItemId: number;
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
        option: '-f, --folder [folder]'
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
      'contentType',
      'folder'
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['listId', 'listTitle', 'listUrl']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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

      let contentTypeName: string = '';
      let targetFolderServerRelativeUrl: string = '';

      if (this.verbose) {
        logger.logToStderr(`Getting content types for list ${args.options.listId || args.options.listTitle || args.options.listUrl}...`);
      }

      let requestOptions: any = {
        url: `${requestUrl}/contenttypes?$select=Name,Id`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const ctypes = await request.get<any>(requestOptions);
      if (args.options.contentType) {
        const foundContentType = ctypes.value.filter((ct: any) => {
          const contentTypeMatch: boolean = ct.Id.StringValue === args.options.contentType || ct.Name === args.options.contentType;

          if (this.debug) {
            logger.logToStderr(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
          }

          return contentTypeMatch;
        });

        if (this.debug) {
          logger.logToStderr('content type filter output...');
          logger.logToStderr(foundContentType);
        }

        if (foundContentType.length > 0) {
          contentTypeName = foundContentType[0].Name;
        }

        // After checking for content types, throw an error if the name is blank
        if (!contentTypeName || contentTypeName === '') {
          throw `Specified content type '${args.options.contentType}' doesn't exist on the target list`;
        }

        if (this.debug) {
          logger.logToStderr(`using content type name: ${contentTypeName}`);
        }
      }

      if (args.options.folder) {
        if (this.debug) {
          logger.logToStderr('setting up folder lookup response ...');
        }

        requestOptions = {
          url: `${requestUrl}/rootFolder`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const rootFolderResponse = await request.get<any>(requestOptions);
        targetFolderServerRelativeUrl = urlUtil.getServerRelativePath(rootFolderResponse["ServerRelativeUrl"], args.options.folder as string);
        await spo.ensureFolder(args.options.webUrl, targetFolderServerRelativeUrl, logger, this.debug);
      }

      if (this.verbose) {
        logger.logToStderr(`Creating item in list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site ${args.options.webUrl}...`);
      }

      const requestBody: any = {
        formValues: this.mapRequestBody(args.options)
      };

      if (args.options.folder) {
        requestBody.listItemCreateInfo = {
          FolderPath: {
            DecodedUrl: targetFolderServerRelativeUrl
          }
        };
      }

      if (args.options.contentType && contentTypeName !== '') {
        if (this.debug) {
          logger.logToStderr(`Specifying content type name [${contentTypeName}] in request body`);
        }

        requestBody.formValues.push({
          FieldName: 'ContentType',
          FieldValue: contentTypeName
        });
      }

      requestOptions = {
        url: `${requestUrl}/AddValidateUpdateItemUsingPath()`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const response = await request.post<any>(requestOptions);

      // Response is from /AddValidateUpdateItemUsingPath POST call, perform get on added item to get all field values
      const fieldValues: FieldValue[] = response.value;
      const idField = fieldValues.filter((thisField) => {
        return (thisField.FieldName === "Id");
      });

      if (this.debug) {
        logger.logToStderr(`field values returned:`);
        logger.logToStderr(fieldValues);
        logger.logToStderr(`Id returned by AddValidateUpdateItemUsingPath: ${idField}`);
      }

      if (idField.length === 0) {
        throw `Item didn't add successfully`;
      }

      requestOptions = {
        url: `${requestUrl}/items(${idField[0].FieldValue})`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const item = await request.get(requestOptions);
      logger.log(<ListItemInstance>item);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = [];
    const excludeOptions: string[] = [
      'listTitle',
      'listId',
      'webUrl',
      'contentType',
      'folder',
      'debug',
      'verbose',
      'output'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody.push({ FieldName: key, FieldValue: `${(<any>options)[key]}` });
      }
    });
    return requestBody;
  }
}

module.exports = new SpoListItemAddCommand();