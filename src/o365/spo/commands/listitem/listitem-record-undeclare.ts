import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import { ContextInfo, ClientSvcResponseContents, ClientSvcResponse } from '../../spo';


const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  id: string;

}

class SpoListItemRecordUndeclareCommand extends SpoCommand {
  private objectIdentity: string;
  constructor() {
    super()/* istanbul ignore next */;
    this.objectIdentity = '';
  }
  public get name(): string {
    return commands.LISTITEM_RECORD_UNDECLARE;
  }

  public get description(): string {
    return 'Undeclares  listitem  as a record';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    let siteAccessToken: string = '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);


    let formDigestValue: string = '';
    let environmentListId: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise | Promise<any> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
          cmd.log(``);
          cmd.log(`auth object:`);
          cmd.log(auth);
        }
        if (this.verbose) {
          cmd.log(`Getting list id...`);
        }

          const listRequestOptions: any = {
            url: `${listRestUrl}/id`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'accept': 'application/json;odata=nometadata'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request for list id...');
            cmd.log(listRequestOptions);
            cmd.log('');
          }
          return request.get(listRequestOptions)
      
      })
      .then((dataReturned: any): request.RequestPromise | Promise<void> => {
        if (dataReturned) {
          environmentListId = dataReturned.value;
          if (this.debug) {
            cmd.log(`data returned[0]:`);
            cmd.log(dataReturned);
            cmd.log(`Retrieved list id ${environmentListId}.`);
          }
        }
        if (this.debug) {
          cmd.log(`getting request digest for request`);
        }
        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }
        if (this.verbose) {
          cmd.log(`Undeclare list item  as a record in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }
        formDigestValue = res['FormDigestValue'];
        return this.requestObjectIdentity(args.options.webUrl, cmd, formDigestValue, siteAccessToken);
      }).then((objectIdentity: string): request.RequestPromise => {
        this.objectIdentity = objectIdentity;
        const requestBody: any = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"
                                  xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                                  <Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}"
                                  Name="UndeclareItemAsRecord" Id="53"><Parameters><Parameter ObjectPathId="49" /></Parameters>
                                  </StaticMethod></Actions><ObjectPaths><Identity Id="49" Name="${this.objectIdentity}:list:${environmentListId}:item:${args.options.id},1" /></ObjectPaths></Request>`
        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'Content-Type': 'text/xml',
            'X-RequestDigest': formDigestValue,
          }),
          body: requestBody
        }

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
          cmd.log('Body:');
          cmd.log(requestBody);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((): void => {
        // REST post call doesn't return anything
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list item should be undeclared as a record'
      },
      {
        option: '-i, --id <id>',
        description: 'ID of the list item to be undeclared as a record.'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list where the list item should be undeclared as a record. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list where the list item should be undeclared as a record. Specify listId or listTitle but not both'
      }

    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public types(): CommandTypes {
    return {
      string: [
        'webUrl',
        'listId',
        'listTitle',
        'id'
      ]
    };
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      const id: number = parseInt(args.options.id);
      if (isNaN(id)) {
        return `${args.options.id} is not a valid list item ID`;
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

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To undeclare an item as a record in a list, you have to first log in to SharePoint using
    the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  Undeclare the list item as a record with ID ${chalk.grey(1)} from list with ID
  ${chalk.grey('0cd891ef-afce-4e55-b836-fce03286cccf')} located in site
  ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
    ${chalk.grey(config.delimiter)} ${commands.LISTITEM_RECORD_UNDECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 1

    Undeclare the list item as a record with ID ${chalk.grey(1)} from list with title
  ${chalk.grey('List 1')} located in site
  ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
    ${chalk.grey(config.delimiter)} ${commands.LISTITEM_RECORD_UNDECLARE} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --id 1
     `);
  }

  /**
   * Requests web object identity for the current web.
   * This request has to be send before we can construct the property bag request.
   * The response data looks like:
   * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * _ObjectType_=SP.Web
   * ServerRelativeUrl=/sites/contoso
   * The ObjectIdentity is needed to create another request to retrieve the property bag or set property.
   * @param webUrl web url
   * @param cmd command cmd
   */
  protected requestObjectIdentity(webUrl: string, cmd: CommandInstance, formDigestValue: string, siteAccessToken: string): Promise<string> {
    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        'X-RequestDigest': formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
    };

    return new Promise<string>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any) => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');

          cmd.log('Attempt to get _ObjectIdentity_ key values');
        }

        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const identityObject = json.find(x => { return x['_ObjectIdentity_'] });
        if (identityObject) {
          resolve(identityObject['_ObjectIdentity_']);
        }
        reject('Cannot proceed. _ObjectIdentity_ not found');
      }).catch((err) => {
        reject(err);
      });
    });
  }
}

module.exports = new SpoListItemRecordUndeclareCommand();