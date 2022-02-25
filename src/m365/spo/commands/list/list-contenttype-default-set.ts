import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import {
  CommandOption,
  CommandTypes
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
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

  public types(): CommandTypes | undefined {
    return {
      string: ['contentTypeId', 'c']
    };
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const baseUrl: string = args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')` :
      `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;

    if (this.verbose) {
      logger.logToStderr('Retrieving content type order...');
    }

    this
      .getContentTypeOrder(baseUrl, logger)
      .then((contentTypeOrder: StringValue[]): Promise<void> => {
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
            // content type is already set as default. we're done
            return Promise.resolve();
          }

          if (this.verbose) {
            logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
          }

          // remove content type from the order array so that we can put it at
          // the beginning to make it default content type          
          contentTypeOrder.splice(contentTypeIndex, 1);
          contentTypeOrder.unshift({
            StringValue: args.options.contentTypeId
          });

          return this.updateContentTypeOrder(baseUrl, contentTypeOrder);
        }

        if (this.debug) {
          logger.logToStderr(`Content type ${args.options.contentTypeId} is not visible in the list`);
        }

        if (this.verbose) {
          logger.logToStderr('Retrieving list content types...');
        }

        // content type is not visible...
        // check if content type exists in the list
        return this
          .getListContentTypes(baseUrl)
          .then((contentTypes: string[]): Promise<void> => {
            if (!contentTypes.find(ct => ct.toUpperCase() === args.options.contentTypeId.toUpperCase())) {
              return Promise.reject(`Content type ${args.options.contentTypeId} missing in the list. Add the content type to the list first and try again.`);
            }

            if (this.verbose) {
              logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
            }

            contentTypeOrder.unshift({
              StringValue: args.options.contentTypeId
            });

            return this.updateContentTypeOrder(baseUrl, contentTypeOrder);
          }, err => Promise.reject(err));
      })
      .then(res => {
        logger.log(res);
        cb();
      }, err => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getContentTypeOrder(baseUrl: string, logger: Logger): Promise<StringValue[]> {
    const requestOptions: AxiosRequestConfig = {
      url: `${baseUrl}/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<{ ContentTypeOrder: StringValue[]; UniqueContentTypeOrder: StringValue[] | null }>(requestOptions)
      .then(contentTypeOrder => {
        let uniqueContentTypeOrder = contentTypeOrder.ContentTypeOrder;
        if (contentTypeOrder.UniqueContentTypeOrder !== null) {
          if (this.debug) {
            logger.logToStderr('Using unique content type order');
          }
          uniqueContentTypeOrder = contentTypeOrder.UniqueContentTypeOrder as StringValue[];
        }
        else {
          if (this.debug) {
            logger.logToStderr('Unique content type order not defined. Using content type order');
          }
        }

        return Promise.resolve(uniqueContentTypeOrder);
      }, err => Promise.reject(err));
  }

  private updateContentTypeOrder(baseUrl: string, contentTypeOrder: StringValue[]): Promise<void> {
    const requestOptions: AxiosRequestConfig = {
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

    return request.post(requestOptions);
  }

  private getListContentTypes(baseUrl: string): Promise<string[]> {
    const requestOptions: AxiosRequestConfig = {
      url: `${baseUrl}/ContentTypes?$select=Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<ContentTypes>(requestOptions)
      .then(contentTypes => contentTypes.value.map(ct => ct.Id.StringValue),
        err => Promise.reject(err));
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
        option: '-c, --contentTypeId <contentTypeId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.listId) {
      if (!validation.isValidGuid(args.options.listId)) {
        return `${args.options.listId} is not a valid GUID`;
      }
    }

    if (args.options.listId && args.options.listTitle) {
      return 'Specify listId or listTitle, but not both';
    }

    if (!args.options.listId && !args.options.listTitle) {
      return 'Specify listId or listTitle, one is required';
    }

    return true;
  }
}

module.exports = new SpoListContentTypeDefaultSetCommand();