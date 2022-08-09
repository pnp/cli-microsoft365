import { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, IdentityResponse, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listId?: string;
  listTitle?: string;
  webUrl: string;
}

export interface ListItemIsRecord {
  CallInfo: any;
  CallObjectId: any;
  IsRecord: any;
}

class SpoListItemIsRecordCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ISRECORD;
  }

  public get description(): string {
    return 'Checks if the specified list item is a record';
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
        listTitle: typeof args.options.listTitle !== 'undefined'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const id: number = parseInt(args.options.id);
        if (isNaN(id)) {
          return `${args.options.id} is not a valid list item ID`;
        }

        if (id < 1) {
          return `Item ID must be a positive number`;
        }

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

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const listIdArgument: string = args.options.listId || '';
    const listTitleArgument: string = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitleArgument)}')`);

    let formDigestValue: string = '';
    let listId: string = '';

    if (this.debug) {
      logger.logToStderr(`Retrieving access token for ${resource}...`);
    }

    ((): Promise<{ Id: string; }> => {
      if (typeof args.options.listId !== 'undefined') {
        if (this.verbose) {
          logger.logToStderr(`List Id passed in as an argument.`);
        }

        return Promise.resolve({ Id: args.options.listId });
      }

      if (this.verbose) {
        logger.logToStderr(`Getting list id...`);
      }
      const requestOptions: any = {
        url: `${listRestUrl}?$select=Id`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      return request.get(requestOptions);
    })()
      .then((res: { Id: string }): Promise<ContextInfo> => {
        listId = res.Id;

        if (this.debug) {
          logger.logToStderr(`Getting request digest for request`);
        }

        return spo.getRequestDigest(args.options.webUrl);
      })
      .then((res: ContextInfo): Promise<IdentityResponse> => {
        formDigestValue = res.FormDigestValue;
        return spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      })
      .then((webIdentityResp: IdentityResponse): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Checking if list item is a record in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }

        const requestBody = this.getIsRecordRequestBody(webIdentityResp.objectIdentity, listId, args.options.id);
        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'Content-Type': 'text/xml',
            'X-RequestDigest': formDigestValue
          },
          data: requestBody
        };

        return request.post<string>(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];

        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          const result: boolean = json[json.length - 1];
          logger.log(result);
          cb();
        }
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private getIsRecordRequestBody(webIdentity: string, listId: string, id: string): string {
    const requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
            <Actions>
              <StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="IsRecord" Id="1"><Parameters><Parameter ObjectPathId="14" /></Parameters></StaticMethod>
            </Actions>
            <ObjectPaths>
              <Identity Id="14" Name="${webIdentity}:list:${listId}:item:${id},1" />
            </ObjectPaths>
          </Request>`;

    return requestBody;
  }
}

module.exports = new SpoListItemIsRecordCommand();