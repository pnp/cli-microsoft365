import { Logger } from '../../../../cli';
import {
  CommandOption,
  CommandTypes
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderExtensions } from '../../FolderExtensions';
import { ListItemInstance } from './ListItemInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.contentType = typeof args.options.contentType !== 'undefined';
    telemetryProps.folder = typeof args.options.folder !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);
    let contentTypeName: string = '';
    let targetFolderServerRelativeUrl: string = '';
    const folderExtensions: FolderExtensions = new FolderExtensions(logger, this.debug);

    if (this.verbose) {
      logger.logToStderr(`Getting content types for list...`);
    }

    const requestOptions: any = {
      url: `${listRestUrl}/contenttypes?$select=Name,Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((response: any): Promise<void> => {
        if (args.options.contentType) {
          const foundContentType = response.value.filter((ct: any) => {
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
            return Promise.reject(`Specified content type '${args.options.contentType}' doesn't exist on the target list`);
          }

          if (this.debug) {
            logger.logToStderr(`using content type name: ${contentTypeName}`);
          }
        }

        if (args.options.folder) {
          if (this.debug) {
            logger.logToStderr('setting up folder lookup response ...');
          }

          const requestOptions: any = {
            url: `${listRestUrl}/rootFolder`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          return request
            .get<any>(requestOptions)
            .then(rootFolderResponse => {
              targetFolderServerRelativeUrl = Utils.getServerRelativePath(rootFolderResponse["ServerRelativeUrl"], args.options.folder as string);

              return folderExtensions.ensureFolder(args.options.webUrl, targetFolderServerRelativeUrl);
            });
        }
        else {
          return Promise.resolve();
        }
      })
      .then((): Promise<any> => {
        if (this.verbose) {
          logger.logToStderr(`Creating item in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
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

        const requestOptions: any = {
          url: `${listRestUrl}/AddValidateUpdateItemUsingPath()`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          data: requestBody,
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((response: any): Promise<any> => {
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
          return Promise.reject(`Item didn't add successfully`);
        }

        const requestOptions: any = {
          url: `${listRestUrl}/items(${idField[0].FieldValue})`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((response: any): void => {
        logger.log(<ListItemInstance>response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
        option: '-c, --contentType [contentType]'
      },
      {
        option: '-f, --folder [folder]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public types(): CommandTypes {
    return {
      string: [
        'webUrl',
        'listId',
        'listTitle',
        'contentType',
        'folder'
      ]
    };
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
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
      !Utils.isValidGuid(args.options.listId)) {
      return `${args.options.listId} in option listId is not a valid GUID`;
    }

    return true;
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