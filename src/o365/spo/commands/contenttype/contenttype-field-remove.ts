import auth from '../../SpoAuth';
// import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
// import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, // CommandValidate, CommandTypes,
  CommandTypes,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
// import { FieldLink } from './FieldLink';
// REST URL
// POST http://<sitecollection>/<site>/_api/web/lists(listid)/contenttypes(contenttypeidcontenttypeid})/fieldlinks(fieldlinkid)/deleteObject()

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
}

class SpoContentTypeFieldRemoveCommand extends SpoCommand {
  private requestDigest: string = '';
  private webId: string = '';
  private siteId: string = '';
  private siteAccessToken: string = '';
  private listId: string = '';

  public get name(): string {
    return `${commands.CONTENTTYPE_FIELD_REMOVE}`;
  }

  public get description(): string {
    return 'Remove a site column reference from a site or list content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['i', 'contentTypeId', 'f', 'fieldLinkId']
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        this.siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Get SiteId required by ProcessQuery endpoint.`);
        }

        // GET SiteId
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/site?$select=Id`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${this.siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        }

        if (this.debug) {
          cmd.log('Executing web request:');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: { Id: string }) => {
        this.siteId = res.Id;

        if (this.debug) {
          cmd.log(`SiteId: ${this.siteId}`);
          cmd.log(`Get WebId required by ProcessQuery endpoint.`);
        }

        // GET WebId
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web?$select=Id`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${this.siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        }

        if (this.debug) {
          cmd.log('Executing web request:');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: { Id: string }) => {
        this.webId = res.Id;

        if (this.debug) {
          cmd.log(`WebId: ${this.webId}`);
        }

        // If ListTitle is provided
        if (args.options.listTitle) {
          // request for the list title Id
          // GET WebId
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/lists/GetByTitle('${args.options.listTitle}')?$select=Id`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${this.siteAccessToken}`,
              accept: 'application/json;odata=nometadata'
            }),
            json: true
          }

          if (this.debug) {
            cmd.log('Executing list request:');
            cmd.log(requestOptions);
            cmd.log('');
          }
          return request.get(requestOptions);
        } 
        else {
          return Promise.resolve(null);
        }
      })
      .then((res: { Id: string }) => {
        if (res) {
          this.listId = res.Id;

          if (this.debug) {
            cmd.log(`ListId: ${this.listId}`);
          }
        }

        return this.getRequestDigestForSite(args.options.webUrl, this.siteAccessToken, cmd, this.debug)
      })
      .then((res: ContextInfo) => {
        if (this.debug) {
          cmd.log(`Form digest='${res.FormDigestValue}`);
          cmd.log('');
        }
        this.requestDigest = res.FormDigestValue;

        const updateChildContentTypes = args.options.listTitle ? false : args.options.updateChildContentTypes;

        if (this.debug) {
          let additionalLog = args.options.listTitle ? `; ListTitle='${args.options.listTitle}'` : ` ; UpdateChildContentTypes='${updateChildContentTypes}`;
          cmd.log(`Remove FieldLink from ContentType. FieldLinkId='${args.options.fieldLinkId}' ; ContentTypeId='${args.options.contentTypeId}' ${additionalLog}`);
          cmd.log(`Execute ProcessQuery.`);
          cmd.log('');
        }

        let requestBody = '';
        if (this.listId) {
          requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.fieldLinkId}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:list:${this.listId}:contenttype:${args.options.contentTypeId}" /></ObjectPaths></Request>`;
        }
        else {
          requestBody =  `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.fieldLinkId}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${args.options.contentTypeId}" /></ObjectPaths></Request>`
        }
        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${this.siteAccessToken}`,
            'X-RequestDigest': this.requestDigest
          }),
          body: requestBody
        };

        if (this.debug) {
          cmd.log('Executing web request.');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (error: any): void => {
        this.handleRejectedODataJsonPromise(error, cmd, cb);
      });
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.contentTypeId = (!(!args.options.contentTypeId)).toString();
    telemetryProps.fieldLinkId = (!(!args.options.fieldLinkId)).toString();
    return telemetryProps;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the content type is located'
      },
      {
        option: '-l, --listTitle <listTitle>',
        description: 'Title of the list where the content type is located (if it is a list content type)'
      },
      {
        option: '-i, --contentTypeId <id>',
        description: 'The ID of the content type to process'
      },
      {
        option: '-f, --fieldLinkId <id>',
        description: 'The ID of the field to remove'
      },
      {
        option: '-c, --updateChild <updateChildContentTypes>',
        description: 'Update child content types'
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

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site using the
      ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To remove a field link from content type, you have to first log in to a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('contentTypeID')} or ${chalk.grey('fieldLinkID')} doesn't refer to an existing objects, you will get
    a an error.

  Examples:
  
    Remove fieldLink with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')} from content type with ID ${chalk.grey('0x0100CA0FA0F5DAEF784494B9C6020C3020A6')}
    from web with Url ${chalk.grey('https://contoso.sharepoint.com')}
      ${chalk.grey(config.delimiter)} ${this.name}  -i "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" -f "880d2f46-fccb-43ca-9def-f88e722cef80" -u https://contoso.sharepoint.com

`);
  }
}

module.exports = new SpoContentTypeFieldRemoveCommand();