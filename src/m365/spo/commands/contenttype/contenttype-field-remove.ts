import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandTypes, CommandValidate, CommandError } from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';

const vorpal: Vorpal = require('../../../../vorpal-init');
interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentTypeId: string;
  fieldLinkId: string;
  webUrl: string;
  listTitle?: string;
  updateChildContentTypes?: boolean;
  confirm?: boolean;
}

class SpoContentTypeFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CONTENTTYPE_FIELD_REMOVE}`;
  }

  public get description(): string {
    return 'Removes a column from a site- or list content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['i', 'contentTypeId']
    };
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.updateChildContentTypes = (!(!args.options.updateChildContentTypes)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let webId: string = '';
    let siteId: string = '';
    let listId: string = '';

    const removeFieldLink = (): void => {
      if (this.debug) {
        cmd.log(`Get SiteId required by ProcessQuery endpoint.`);
      }

      // GET SiteId
      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/site?$select=Id`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        json: true
      }

      request
        .get<{ Id: string }>(requestOptions)
        .then((res: { Id: string }): Promise<{ Id: string; }> => {
          siteId = res.Id;

          if (this.debug) {
            cmd.log(`SiteId: ${siteId}`);
            cmd.log(`Get WebId required by ProcessQuery endpoint.`);
          }

          // GET WebId
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web?$select=Id`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            json: true
          }

          return request.get(requestOptions);
        })
        .then((res: { Id: string }): Promise<{ Id: string; }> => {
          webId = res.Id;

          if (this.debug) {
            cmd.log(`WebId: ${webId}`);
          }

          // If ListTitle is provided
          if (!args.options.listTitle) {
            return Promise.resolve(undefined as any);
          }
          // Request for the ListId
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/lists/GetByTitle('${encodeURIComponent(args.options.listTitle)}')?$select=Id`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            json: true
          }

          return request.get(requestOptions);
        })
        .then((res?: { Id: string }): Promise<ContextInfo> => {
          if (res) {
            listId = res.Id;

            if (this.debug) {
              cmd.log(`ListId: ${listId}`);
            }
          }

          return this.getRequestDigest(args.options.webUrl)
        })
        .then((res: ContextInfo): Promise<string> => {
          const requestDigest: string = res.FormDigestValue;

          const updateChildContentTypes: boolean = args.options.listTitle ? false : args.options.updateChildContentTypes === true;

          if (this.debug) {
            const additionalLog = args.options.listTitle ? `; ListTitle='${args.options.listTitle}'` : ` ; UpdateChildContentTypes='${updateChildContentTypes}`;
            cmd.log(`Remove FieldLink from ContentType. FieldLinkId='${args.options.fieldLinkId}' ; ContentTypeId='${args.options.contentTypeId}' ${additionalLog}`);
            cmd.log(`Execute ProcessQuery.`);
            cmd.log('');
          }

          let requestBody: string = '';
          if (listId) {
            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${Utils.escapeXml(args.options.fieldLinkId)}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${Utils.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
          }
          else {
            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${Utils.escapeXml(args.options.fieldLinkId)}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${Utils.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`
          }

          const requestOptions: any = {
            url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': requestDigest
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
            return;
          }
          if (this.debug) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
          cb();
        }, (error: any): void => {
          this.handleRejectedODataJsonPromise(error, cmd, cb);
        });
    }

    if (args.options.confirm) {
      removeFieldLink();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the column ${args.options.fieldLinkId} from content type ${args.options.contentTypeId}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeFieldLink();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the content type is located'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the content type is located (if it is a list content type)'
      },
      {
        option: '-i, --contentTypeId <contentTypeId>',
        description: 'The ID of the content type to remove the column from'
      },
      {
        option: '-f, --fieldLinkId <fieldLinkId>',
        description: 'The ID of the column to remove'
      },
      {
        option: '-c, --updateChildContentTypes',
        description: 'Update child content types'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removal of a column from content type'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.contentTypeId) {
        return 'Required parameter contentTypeId missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (!args.options.fieldLinkId) {
        return 'Required parameter fieldLinkId missing';
      }

      if (!Utils.isValidGuid(args.options.fieldLinkId)) {
        return `${args.options.fieldLinkId} is not a valid GUID`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Remove column with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')} from
    content type with ID ${chalk.grey('0x0100CA0FA0F5DAEF784494B9C6020C3020A6')}
    from web with URL ${chalk.grey('https://contoso.sharepoint.com')}
      ${this.name} --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --confirm

    Remove column with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')} from
    content type with ID ${chalk.grey('0x0100CA0FA0F5DAEF784494B9C6020C3020A6')}
    from web with URL ${chalk.grey('https://contoso.sharepoint.com')} updating child content types
      ${this.name} --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --updateChildContentTypes 

    Remove fieldLink with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')} from list
    content type with ID ${chalk.grey('0x0100CA0FA0F5DAEF784494B9C6020C3020A6')}
    from web with URL ${chalk.grey('https://contoso.sharepoint.com')} 
      ${this.name} --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A60062F089A38C867747942DB2C3FC50FF6A" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --listTitle "Documents"
      `);
  }
}

module.exports = new SpoContentTypeFieldRemoveCommand();