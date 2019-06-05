import config from "../../../../config";
import commands from "../../commands";
import GlobalOptions from "../../../../GlobalOptions";
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from "../../../../Command";
import SpoCommand from "../../../base/SpoCommand";
import Utils from "../../../../Utils";
import {
  ContextInfo,
  ClientSvcResponse,
  ClientSvcResponseContents,
} from "../../spo";
import { ClientSvc, IdentityResponse } from "../../ClientSvc";

const vorpal: Vorpal = require("../../../../vorpal-init");

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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.date = typeof args.options.date !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const clientSvc: ClientSvc = new ClientSvc(cmd, this.debug);
    let formDigestValue: string = '';
    let webIdentity: string = '';
    let listId: string = '';

    const listRestUrl: string = args.options.listId
      ? `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;

    this
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        formDigestValue = contextResponse.FormDigestValue;

        return clientSvc.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
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
          json: true
        }

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
          body: requestBody
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
          cmd.log(result);
          cb();
        }
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the list is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'The ID of the list where the item is located. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'The title of the list where the item is located. Specify listId or listTitle but not both'
      },
      {
        option: '-i, --id <id>',
        description: 'The ID of the list item to declare as record'
      },
      {
        option: '-d, --date [date]',
        description: 'Record declaration date in ISO format, eg. 2019-12-31'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

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

      if (args.options.listId && !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      if (!args.options.id) {
        return `Specify id for item to declare as a record`;
      }

      const id: number = parseInt(args.options.id);
      if (isNaN(id)) {
        return `${args.options.id} is not a number`;
      }

      if (id < 1) {
        return `Item ID must be a positive number`;
      }

      if (args.options.date && !Utils.isValidISODate(args.options.date)) {
        return `${args.options.date} in option date is not in ISO format (yyyy-mm-dd)`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Declare a document with id ${chalk.grey("1")} as a record in list with title ${chalk.grey("Demo List")}
    located in site ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --id 1

    Declare a document with id ${chalk.grey("1")} as a record in list with id
    ${chalk.grey("ea8e1109-2013-1a69-bc05-1403201257fc")} located in site
    ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listId ea8e1109-2013-1a69-bc05-1403201257fc --id 1
  
    Declare a document with id ${chalk.grey("1")} as a record with record declaration date
    ${chalk.grey("March 14, 2012")} in list with title ${chalk.grey("Demo List")} located in site
    ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --id 1 --date 2012-03-14

    Declare a document with id ${chalk.grey("1")} as a record with record declaration date
    ${chalk.grey("September 3, 2013")} in list with id ${chalk.grey("ea8e1356-5910-abc9-bc05-2408198057fc")}
    located in site ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listId ea8e1356-5910-abc9-bc05-2408198057fc --id 1 --date 2013-09-03
   `
    );
  }
}
module.exports = new SpoListItemRecordDeclareCommand();