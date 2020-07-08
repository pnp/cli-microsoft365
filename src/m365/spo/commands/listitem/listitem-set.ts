import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ListItemInstance } from './ListItemInstance';
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
  contentType?: string;
  systemUpdate?: boolean;
}

class SpoListItemSetCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_SET;
  }

  public get description(): string {
    return 'Updates a list item in the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.contentType = typeof args.options.contentType !== 'undefined';
    telemetryProps.systemUpdate = typeof args.options.systemUpdate !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);
    let contentTypeName: string = '';

    let formDigestValue: string = '';
    let environmentListId: string = '';

    ((): Promise<any> => {
      if (args.options.systemUpdate) {
        if (this.verbose) {
          cmd.log(`Getting list id...`);
        }

        const listRequestOptions: any = {
          url: `${listRestUrl}/id`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(listRequestOptions)
      }
      else {
        return Promise.resolve();
      }
    })()
      .then((dataReturned: any): Promise<void> => {
        if (dataReturned) {
          environmentListId = dataReturned.value;
        }

        if (args.options.contentType) {
          if (this.verbose) {
            cmd.log(`Getting content types for list...`);
          }

          const requestOptions: any = {
            url: `${listRestUrl}/contenttypes?$select=Name,Id`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            json: true
          };

          return request.get(requestOptions);
        }
        else {
          return Promise.resolve();
        }
      })
      .then((response: any): Promise<ContextInfo> => {
        if (args.options.contentType) {
          if (this.debug) {
            cmd.log('content type lookup response...');
            cmd.log(response);
          }

          const foundContentType: { Name: string; }[] = response.value.filter((ct: any) => {
            const contentTypeMatch: boolean = ct.Id.StringValue === args.options.contentType || ct.Name === args.options.contentType;

            if (this.debug) {
              cmd.log(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
            }

            return contentTypeMatch;
          });

          if (this.debug) {
            cmd.log('content type filter output...');
            cmd.log(foundContentType);
          }

          if (foundContentType.length > 0) {
            contentTypeName = foundContentType[0].Name;
          }

          // After checking for content types, throw an error if the name is blank
          if (!contentTypeName || contentTypeName === '') {
            return Promise.reject(`Specified content type '${args.options.contentType}' doesn't exist on the target list`);
          }

          if (this.debug) {
            cmd.log(`using content type name: ${contentTypeName}`);
          }
        }
        if (args.options.systemUpdate) {
          if (this.debug) {
            cmd.log(`getting request digest for systemUpdate request`);
          }

          return this.getRequestDigest(args.options.webUrl);
        }
        else {
          return Promise.resolve(undefined as any);
        }
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Updating item in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }

        formDigestValue = args.options.systemUpdate ? res['FormDigestValue'] : '';

        if (args.options.systemUpdate) {
          return this.requestObjectIdentity(args.options.webUrl, cmd, formDigestValue);
        }

        return Promise.resolve('');
      }).then((objectIdentity: string): Promise<any> => {
        const additionalContentType: string = (args.options.systemUpdate && args.options.contentType && contentTypeName !== '') ? `
              <Parameters>
                <Parameter Type="String">ContentType</Parameter>
                <Parameter Type="String">${contentTypeName}</Parameter>
              </Parameters>`
          : ``;

        const requestBody: any = args.options.systemUpdate ?
          `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
            <Actions>
              <Method Name="ParseAndSetFieldValue" Id="1" ObjectPathId="147">${this.mapRequestBody(args.options).join()}${additionalContentType}
              </Method>
              <Method Name="SystemUpdate" Id="2" ObjectPathId="147" />
            </Actions>
            <ObjectPaths>
              <Identity Id="147" Name="${objectIdentity}:list:${environmentListId}:item:${args.options.id},1" />
            </ObjectPaths>
          </Request>`
          : {
            formValues: this.mapRequestBody(args.options)
          };

        if (args.options.contentType && contentTypeName !== '' && !args.options.systemUpdate) {
          if (this.debug) {
            cmd.log(`Specifying content type name [${contentTypeName}] in request body`);
          }

          requestBody.formValues.push({
            FieldName: 'ContentType',
            FieldValue: contentTypeName
          });
        }

        const requestOptions: any = args.options.systemUpdate ? {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'Content-Type': 'text/xml',
            'X-RequestDigest': formDigestValue,
          },
          body: requestBody
        } : {
            url: `${listRestUrl}/items(${args.options.id})/ValidateUpdateListItem()`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            body: requestBody,
            json: true
          };

        return request.post(requestOptions);
      })
      .then((response: any): Promise<any> => {
        let itemId: number = 0;

        if (args.options.systemUpdate) {
          if (response.indexOf("ErrorMessage") > -1) {
            return Promise.reject(`Error occurred in systemUpdate operation - ${response}`);
          }
          else {
            itemId = Number(args.options.id);
          }
        }
        else {
          // Response is from /ValidateUpdateListItem POST call, perform get on updated item to get all field values
          const returnedData: any = response.value;

          if (!returnedData[0].ItemId) {
            return Promise.reject(`Item didn't update successfully`)
          }
          else {
            itemId = returnedData[0].ItemId
          }
        }

        const requestOptions: any = {
          url: `${listRestUrl}/items(${itemId})`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((response: any): void => {
        cmd.log(<ListItemInstance>response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the item should be updated'
      },
      {
        option: '-i, --id <id>',
        description: 'ID of the list item to update.'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list where the item should be updated. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list where the item should be updated. Specify listId or listTitle but not both'
      },
      {
        option: '-c, --contentType [contentType]',
        description: 'The name or the ID of the content type to associate with the updated item'
      },
      {
        option: '-s, --systemUpdate',
        description: 'Update the item without updating the modified date and modified by fields'
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
        'id',
        'contentType',
      ],
      boolean: [
        'systemUpdate'
      ]
    };
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
      `  Examples:
  
    Update the item with ID of ${chalk.grey('147')} with Title ${chalk.grey('Demo Item')} and content type name
    ${chalk.grey('Item')} in list with title ${chalk.grey('Demo List')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_SET} --contentType Item --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"

    Update an item with Title ${chalk.grey('Demo Multi Managed Metadata Field')} and
    a single-select metadata field named ${chalk.grey('SingleMetadataField')} in list with
    title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_SET} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x  --id 147 --Title "Demo Single Managed Metadata Field" --SingleMetadataField "TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;"

    Update an item with ID of ${chalk.grey('147')} with Title ${chalk.grey('Demo Multi Managed Metadata Field')}
    and a multi-select metadata field named ${chalk.grey('MultiMetadataField')} in list
    with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_SET} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --id 147 --Title "Demo Multi Managed Metadata Field" --MultiMetadataField "TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;"
  
    Update an item with ID of ${chalk.grey('147')} with Title ${chalk.grey('Demo Single Person Field')}
    and a single-select people field named ${chalk.grey('SinglePeopleField')} in list
    with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_SET} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --id 147 --Title "Demo Single Person Field" --SinglePeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'}]"
      
    Update an item with ID of ${chalk.grey('147')} with Title ${chalk.grey('Demo Multi Person Field')}
    and a multi-select people field named ${chalk.grey('MultiPeopleField')} in list
    with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_SET} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --id 147 --Title "Demo Multi Person Field" --MultiPeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'},{'Key':'i:0#.f|membership|adamb@conotoso.com'}]"
    
    Update an item with ID of ${chalk.grey('147')} with Title ${chalk.grey('Demo Hyperlink Field')}
    and a hyperlink field named ${chalk.grey('CustomHyperlink')} in list
    with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_SET} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --id 147 --Title "Demo Hyperlink Field" --CustomHyperlink "https://www.bing.com, Bing"
   `);
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = [];
    const excludeOptions: string[] = [
      'listTitle',
      'listId',
      'webUrl',
      'id',
      'contentType',
      'systemUpdate',
      'debug',
      'verbose',
      'output'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        if (options.systemUpdate) {
          requestBody.push(`
            <Parameters>
              <Parameter Type="String">${key}</Parameter>
              <Parameter Type="String">${(<any>options)[key]}</Parameter>
            </Parameters>`);
        }
        else {
          requestBody.push({ FieldName: key, FieldValue: (<any>options)[key] });
        }
      }
    });

    return requestBody;
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
  private requestObjectIdentity(webUrl: string, cmd: CommandInstance, formDigestValue: string): Promise<string> {
    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestValue
      },
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
    };

    return new Promise<string>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any) => {
        if (this.debug) {
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

        reject('Cannot proceed. _ObjectIdentity_ not found'); // this is not supposed to happen
      }).catch((err) => {
        reject(err);
      });
    });
  }
}

module.exports = new SpoListItemSetCommand();