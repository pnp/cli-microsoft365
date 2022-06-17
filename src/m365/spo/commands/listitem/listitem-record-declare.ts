import { Logger } from '../../../../cli';
import {
  CommandError
} from "../../../../Command";
import config from "../../../../config";
import GlobalOptions from "../../../../GlobalOptions";
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, IdentityResponse, spo, validation } from "../../../../utils";
import SpoCommand from "../../../base/SpoCommand";
import commands from "../../commands";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  date?: string;
  id: string;
  listId?: string;
  listTitle?: string;
  webUrl: string;
}

class SpoListItemRecordDeclareCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RECORD_DECLARE;
  }

  public get description(): string {
    return "Declares the specified list item as a record";
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
        listTitle: typeof args.options.listTitle !== 'undefined',
        date: typeof args.options.date !== 'undefined'
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
        option: '-i, --id <id>'
      },
      {
        option: '-d, --date [date]'
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

        if (!args.options.listId && !args.options.listTitle) {
          return `Specify listId or listTitle`;
        }

        if (args.options.listId && args.options.listTitle) {
          return `Specify listId or listTitle but not both`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        const id: number = parseInt(args.options.id);
        if (isNaN(id)) {
          return `${args.options.id} is not a number`;
        }

        if (id < 1) {
          return `Item ID must be a positive number`;
        }

        if (args.options.date && !validation.isValidISODate(args.options.date)) {
          return `${args.options.date} in option date is not in ISO format (yyyy-mm-dd)`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let formDigestValue: string = '';
    let webIdentity: string = '';
    let listId: string = '';

    const listRestUrl: string = args.options.listId
      ? `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;

    spo
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        formDigestValue = contextResponse.FormDigestValue;

        return spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      })
      .then((webIdentityResp: IdentityResponse): Promise<{ Id: string }> => {
        webIdentity = webIdentityResp.objectIdentity;

        if (args.options.listId) {
          return Promise.resolve({ Id: args.options.listId });
        }

        const requestOptions: any = {
          url: `${listRestUrl}?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: { Id: string }): Promise<string> => {
        listId = res.Id;
        const requestBody: string = this.getDeclareRecordRequestBody(webIdentity, listId, args.options.id, args.options.date || '');

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'Content-Type': 'text/xml',
            'X-RequestDigest': formDigestValue
          },
          data: requestBody
        };

        return request.post(requestOptions);
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

  protected getDeclareRecordRequestBody(webIdentity: string, listId: string, id: string, date: string): string {
    let requestBody: string = '';
    if (date.length === 10) {
      requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="DeclareItemAsRecordWithDeclarationDate" Id="48"><Parameters><Parameter ObjectPathId="21" /><Parameter Type="DateTime">${date}</Parameter></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="21" Name="${webIdentity}:list:${listId}:item:${id},1" /></ObjectPaths></Request>`;
    }
    else {
      requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="DeclareItemAsRecord" Id="37"><Parameters><Parameter ObjectPathId="12" /></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="12" Name="${webIdentity}:list:${listId}:item:${id},1" /></ObjectPaths></Request>`;
    }

    return requestBody;
  }
}
module.exports = new SpoListItemRecordDeclareCommand();