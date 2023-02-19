import { Logger } from '../../../../cli/Logger';
import config from "../../../../config";
import GlobalOptions from "../../../../GlobalOptions";
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from "../../../../utils/formatting";
import { ClientSvcResponse, ClientSvcResponseContents, spo } from "../../../../utils/spo";
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from "../../../../utils/validation";
import SpoCommand from "../../../base/SpoCommand";
import commands from "../../commands";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  date?: string;
  listItemId: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
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
        option: '--listUrl [listUrl]'
      }
      ,
      {
        option: '-i, --listItemId <listItemId>'
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

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        const id: number = parseInt(args.options.listItemId);
        if (isNaN(id)) {
          return `${args.options.listItemId} is not a number`;
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

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let listId: string = '';

      if (args.options.listId) {
        listId = args.options.listId;
      }
      else {
        let requestUrl = `${args.options.webUrl}/_api/web`;

        if (args.options.listTitle) {
          requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
        }
        else if (args.options.listUrl) {
          const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
          requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }

        const requestOptions: CliRequestOptions = {
          url: `${requestUrl}?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const list = await request.get<{ Id: string; }>(requestOptions);
        listId = list.Id;
      }

      const contextResponse = await spo.getRequestDigest(args.options.webUrl);
      const formDigestValue = contextResponse.FormDigestValue;

      const webIdentityResp = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      const webIdentity = webIdentityResp.objectIdentity;
      const requestBody: string = this.getDeclareRecordRequestBody(webIdentity, listId, args.options.listItemId, args.options.date || '');

      const requestOptions: CliRequestOptions = {
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