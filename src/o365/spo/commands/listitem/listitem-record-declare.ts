import auth from "../../SpoAuth";
import config from "../../../../config";
import commands from "../../commands";
import GlobalOptions from "../../../../GlobalOptions";
import * as request from "request-promise-native";
import {
  CommandOption,
  CommandValidate,
  CommandTypes,
  CommandError
} from "../../../../Command";
import SpoCommand from "../../SpoCommand";
import Utils from "../../../../Utils";
import { Auth } from "../../../../Auth";
import {
  ContextInfo,
  ClientSvcResponse,
  ClientSvcResponseContents,
} from "../../spo";
import { ClientSvc, IdentityResponse } from "../../common/ClientSvc";

const vorpal: Vorpal = require("../../../../vorpal-init");

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  id: string;
  date?: string;
}

class SpoListItemDeclareRecord extends SpoCommand {

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_RECORD_DECLARE;
  }

  public get description(): string {
    return "Declares the specified list item as a record";
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== "undefined";
    telemetryProps.listTitle = typeof args.options.listTitle !== "undefined";
    telemetryProps.id = typeof args.options.id !== "undefined";
    telemetryProps.date = typeof args.options.date !== "undefined";
    return telemetryProps;
  }

  public commandAction(
    cmd: CommandInstance,
    args: CommandArgs,
    cb: () => void
  ): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const clientSvc: ClientSvc = new ClientSvc(cmd, this.debug);
    const listIdArgument = args.options.listId || "";
    const listTitleArgument = args.options.listTitle || "";
    const id = args.options.id;
    const dateArgument = args.options.date || "";
    let siteAccessToken: string = "";
    let formDigestValue: string = "";
    let webIdentity: string = "";
    let listId: string = "";


    const listRestUrl: string = args.options.listId
      ? `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(
        listIdArgument
      )}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(
        listTitleArgument
      )}')`;

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then(
        (accessToken: string): request.RequestPromise => {
          siteAccessToken = accessToken;

          if (this.debug) {
            cmd.log(
              `Retrieved access token ${accessToken}. Retrieving request digest...`
            );
          }

          return this.getRequestDigest(cmd, this.debug);
        })
      .then(
        (contextResponse: ContextInfo): Promise<IdentityResponse> => {
          formDigestValue = contextResponse.FormDigestValue;

          if (this.debug) {
            cmd.log("contextResponse:");
            cmd.log(JSON.stringify(contextResponse));
            cmd.log("");
          }

          return clientSvc.getCurrentWebIdentity(
            args.options.webUrl,
            siteAccessToken,
            formDigestValue
          );
        }
      )
      .then((webIdentityResp: IdentityResponse): request.RequestPromise => {
        webIdentity = webIdentityResp.objectIdentity;

        const requestOptions: any = {
          url: `${listRestUrl}?$select=Id`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        }

        if (this.debug) {
          cmd.log('Executing get list web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: { Id: string }): request.RequestPromise => {

        listId = res.Id;
        const requestBody = this.generateDeclareRecordRequestBody(webIdentity, listId, id, dateArgument);

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'Content-Type': 'text/xml',
            'X-RequestDigest': formDigestValue
          }),
          body: requestBody
        };

        if (this.debug) {
          cmd.log("Executing declare item as record web request...");
          cmd.log(requestOptions);
          cmd.log("");
          cmd.log("Body:");
          cmd.log(requestBody);
          cmd.log("");
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];

        if (response.ErrorInfo) {
          cmd.log(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          const result: boolean = json[json.length - 1];
          cmd.log(result);
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  protected generateDeclareRecordRequestBody(webIdentity: string, listId: string, id: string, date: string): string {

    let requestBody = "";
    if (date.length == 10) {
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
        option: "-u, --webUrl <webUrl>",
        description: "URL of the site where the item should be added"
      },
      {
        option: "-l, --listId [listId]",
        description:
          "ID of the list where the item should be added. Specify listId or listTitle but not both"
      },
      {
        option: "-t, --listTitle [listTitle]",
        description:
          "Title of the list where the item should be added. Specify listId or listTitle but not both"
      },
      {
        option: "-i, --id [id]",
        description: "ID of the list item to declare as a record"
      },
      {
        option: "-d, --date [date]",
        description: "Record declaration date"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public types(): CommandTypes {
    return {
      string: ["webUrl", "listId", "listTitle", "id", "date"]
    };
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return "Required parameter webUrl missing";
      }

      const isValidSharePointUrl:
        | boolean
        | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.listId && !args.options.listTitle) {
        return `Specify listId or listTitle`;
      }

      if (args.options.listId && args.options.listTitle) {
        return `Specify listId or listTitle but not both. (${args.options.listId} | ${args.options.listTitle})`;
      }

      if (args.options.listId && !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      if (!args.options.id) {
        return `Specify id for item to declare as a record`;
      }

      if (args.options.id && !Utils.isPositiveInt(args.options.id)) {
        return `Specify id for item to declare as a record`;
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
    log(`  ${chalk.yellow("Important:")} before using this command, 
    log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To declare an item as a record, you have to first log in to SharePoint using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Declare a document with id ${chalk.grey("1")} as a record in list with
    title ${chalk.grey("Demo List")} in site ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --id 1

    Declare a document with id ${chalk.grey("1")} as a record in list with
    id ${chalk.grey("ea8e1109-2013-1a69-bc05-1403201257fc")} in site ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listId ea8e1109-2013-1a69-bc05-1403201257fc --id 1
  
    Declare a document with id ${chalk.grey("1")} as a record with record declaration date ${chalk.grey("15 August 1976")} in list with
    title ${chalk.grey("Demo List")} in site ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --id 1 --date 1976-08-15

    Declare a document with id ${chalk.grey("1")} as a record with record declaration date ${chalk.grey("15 August 1976")} in list with
    id ${chalk.grey("ea8e1356-5910-abc9-bc05-2408198057fc")} in site ${chalk.grey("https://contoso.sharepoint.com/sites/project-x")}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_RECORD_DECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listId ea8e1356-5910-abc9-bc05-2408198057fc --id 1 --date 1976-08-15
   `
    );
  }
}
module.exports = new SpoListItemDeclareRecord();