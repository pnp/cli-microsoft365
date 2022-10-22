import { AxiosRequestConfig } from 'axios';
import { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
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

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['listId', 'listTitle', 'listUrl']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
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

    let formDigestValue: string = '';
    let listId: string = '';

    if (this.debug) {
      logger.logToStderr(`Retrieving access token for ${resource}...`);
    }

    try {
      if (typeof args.options.listId !== 'undefined') {
        if (this.verbose) {
          logger.logToStderr(`List Id passed in as an argument.`);
        }

        listId = args.options.listId;
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Getting list id for list ${args.options.listTitle ? args.options.listTitle : args.options.listId}`);
        }
        const requestOptions: AxiosRequestConfig = {
          url: `${requestUrl}?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const list = await request.get<{ Id: string; }>(requestOptions);
        listId = list.Id;
      }

      if (this.debug) {
        logger.logToStderr(`Getting request digest for request`);
      }

      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      formDigestValue = reqDigest.FormDigestValue;

      const webIdentityResp = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);

      if (this.verbose) {
        logger.logToStderr(`Checking if list item is a record in list ${args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl} in site ${args.options.webUrl}...`);
      }

      const requestBody = this.getIsRecordRequestBody(webIdentityResp.objectIdentity, listId, args.options.id);
      const requestOptions: AxiosRequestConfig = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'Content-Type': 'text/xml',
          'X-RequestDigest': formDigestValue
        },
        data: requestBody
      };

      const res = await request.post<string>(requestOptions);

      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];

      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const result: boolean = json[json.length - 1];
        logger.log(result);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
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