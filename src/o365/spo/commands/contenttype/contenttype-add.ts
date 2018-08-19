import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandTypes, CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listTitle?: string;
  name: string;
  id: string;
  description?: string;
  group?: string;
}

class SpoContentTypeAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CONTENTTYPE_ADD}`;
  }

  public get description(): string {
    return 'Adds a new list or site content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['id', 'i']
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let parentInfo: string = '';

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): Promise<string> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving information about the parent object...`);
        }

        return this.getParentInfo(args.options.listTitle, args.options.webUrl, siteAccessToken, cmd);
      })
      .then((parent: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Parent object');
          cmd.log(parent);
          cmd.log('');
        }

        parentInfo = parent;

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        const description: string = args.options.description ?
          `<Property Name="Description" Type="String">${Utils.escapeXml(args.options.description)}</Property>` :
          '<Property Name="Description" Type="Null" />';
        const group: string = args.options.group ?
          `<Property Name="Group" Type="String">${Utils.escapeXml(args.options.group)}</Property>` :
          '<Property Name="Group" Type="Null" />';
          
        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /></Actions><ObjectPaths><Property Id="7" ParentId="5" Name="ContentTypes" /><Method Id="9" ParentId="7" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}">${description}${group}<Property Name="Id" Type="String">${Utils.escapeXml(args.options.id)}</Property><Property Name="Name" Type="String">${Utils.escapeXml(args.options.name)}</Property><Property Name="ParentContentType" Type="Null" /></Parameter></Parameters></Method>${parentInfo}</ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
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
          return;
        }
        else {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private getParentInfo(listTitle: string | undefined, webUrl: string, siteAccessToken: string, cmd: CommandInstance): Promise<string> {
    return new Promise<string>((resolve: (parentInfo: string) => void, reject: (error: any) => void): void => {
      if (!listTitle) {
        resolve('<Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />');
        return;
      }

      let siteId: string = '';
      let webId: string = '';

      ((): request.RequestPromise => {
        if (this.verbose) {
          cmd.log(`Retrieving site collection id...`);
        }

        const requestOptions: any = {
          url: `${webUrl}/_api/site?$select=Id`,
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

        return request.get(requestOptions);
      })()
        .then((res: { Id: string }): request.RequestPromise => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          siteId = res.Id;

          if (this.verbose) {
            cmd.log(`Retrieving site id...`);
          }

          const requestOptions: any = {
            url: `${webUrl}/_api/web?$select=Id`,
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

          return request.get(requestOptions);
        })
        .then((res: { Id: string }): request.RequestPromise => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          webId = res.Id;

          if (this.verbose) {
            cmd.log(`Retrieving list id...`);
          }

          const requestOptions: any = {
            url: `${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle)}')?$select=Id`,
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

          return request.get(requestOptions);
        })
        .then((res: { Id: string }): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          resolve(`<Identity Id="5" Name="1a48869e-c092-0000-1f61-81ec89809537|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${res.Id}" />`)
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the content type should be created'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the content type should be created (if it should be created as a list content type)'
      },
      {
        option: '-i, --id <id>',
        description: 'The ID of the content type. Determines the parent content type'
      },
      {
        option: '-n, --name <name>',
        description: 'The name of the content type'
      },
      {
        option: '-d, --description [description]',
        description: 'The description of the content type'
      },
      {
        option: '-g, --group [group]',
        description: 'The group with which the content type should be associated'
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

      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To create a content type, you have to first connect to a SharePoint site
    using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If the specified content type already exists, you will get a
    ${chalk.grey('A duplicate content type "Your Content Type" was found.')} error.

    The ID of the content type specifies the parent content type from which this
    content type inherits.

  Examples:
  
    Create a site content type that inherits from the List item content type
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'PnP Alert' --id 0x01007926A45D687BA842B947286090B8F67D --group 'PnP Content Types'
    
    Create a list content type that inherits from the List item content type
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Alerts --name 'PnP Alert' --id 0x01007926A45D687BA842B947286090B8F67D

  More information:

    Content Type IDs
      https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/aa543822(v%3Doffice.14)
`);
  }
}

module.exports = new SpoContentTypeAddCommand();