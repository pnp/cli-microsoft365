import auth from '../../SpoAuth';
// import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
// import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, // CommandValidate, CommandTypes,
  CommandTypes,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
// import { FieldLink } from './FieldLink';
// REST URL
// POST http://<sitecollection>/<site>/_api/web/lists(listid)/contenttypes(contenttypeidcontenttypeid})/fieldlinks(fieldlinkid)/deleteObject()

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}


interface Options extends GlobalOptions {
  contentTypeId: string;
  fieldId: string;
  webUrl: string;
  listTitle?: string;
  updateChildContentTypes?: boolean;
}

class SpoContentTypeFieldRemoveCommand extends SpoCommand {
  private requestDigest: string = "";

  public get name(): string {
    return `${commands.CONTENTTYPE_FIELD_REMOVE}`;
  }

  public get description(): string {
    return 'Remove a site column reference from a site or list content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['i', 'contentTypeId', 'f', 'fieldId']
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let siteId = "";
    let webId = "";

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Get SiteId required by ProcessQuery endpoint.`);
        }

        // GET SiteId
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/site?$select=Id`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
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
        siteId = res.Id;

        if (this.debug) {
          cmd.log(`SiteId: ${siteId}`);
          cmd.log(`Get WebId required by ProcessQuery endpoint.`);
        }

        // GET SiteId
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web?$select=Id`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
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
        webId = res.Id;

        if (this.debug) {
          cmd.log(`WebId: ${siteId}`);

          
        }

        return this.ensureRequestDigest(args.options.webUrl, siteAccessToken, cmd);
      })
      .then(() => {
        const updateChildContentTypes = args.options.listTitle ? args.options.updateChildContentTypes == true : false;

        if (this.debug) {
          let additionalLog = args.options.listTitle ? `; ListTitle='${args.options.listTitle}'` : ` ; UpdateChildContentTypes='${updateChildContentTypes}`;
          cmd.log(`Remove FieldLink from ContentType. FieldLinkId='${args.options.fieldId}' ; ContentTypeId='${args.options.contentTypeId}' ${additionalLog}`);
          cmd.log(`Execute ProcessQuery.`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'X-RequestDigest': this.requestDigest
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.fieldId}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${args.options.contentTypeId}" /></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request.');
          cmd.log(requestOptions);
          cmd.log('');
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
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
          cb();
        }
      }, (error: any): void => {
        if (error === 'DONE') {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
          cb();
        }
        else {
          this.handleRejectedODataJsonPromise(error, cmd, cb);
        }
      });
  }

  /**
   * Apparently - the REST API for removal of the field link doesn't work. Fallback to ProcessQueryx
   * @param cmd 
   * @param args 
   * @param cb 
   */
  public commandActionREST(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Get all field links for ContentType. CTId='${args.options.contentTypeId}' ${args.options.listTitle ? ` ; ListTitle='${args.options.listTitle}'` : ""}`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/` : '')}contenttypes('${encodeURIComponent(args.options.contentTypeId)}')/fieldlinks('${encodeURIComponent(args.options.fieldId)}')/DeleteObject`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        }

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((result: any) => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(result);
          cmd.log('');
        }

        cmd.log(result);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }




  public commandHelp(args: any, log: (message: string) => void): void {
    throw new Error("Method not implemented.");
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
        option: '-f, --fieldId <id>',
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

  private ensureRequestDigest(siteUrl: string, siteAccessToken: string, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this.requestDigest) {
        if (this.debug) {
          cmd.log('Request digest already present');
        }
        resolve();
        return;
      }

      if (this.debug) {
        cmd.log('Retrieving request digest...');
      }

      this
        .getRequestDigestForSite(siteUrl, siteAccessToken, cmd, this.debug)
        .then((res: ContextInfo): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          this.requestDigest = res.FormDigestValue;
          resolve();
        }, (error: any): void => {
          reject(error);
        });
    });
  }
}

module.exports = new SpoContentTypeFieldRemoveCommand();