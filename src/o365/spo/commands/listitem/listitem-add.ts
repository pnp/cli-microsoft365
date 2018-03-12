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
import Auth from '../../../../Auth';
import { ListItemInstance } from './ListItemInstance';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  contentType?: string;
  folder?: string;
}

class SpoListItemAddCommand extends SpoCommand {

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_ADD;
  }

  public get description(): string {
    return 'Creates a list item in the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);

    // TODO: impleemnt getTelementryProperties
    /*
    // add properties with identifiable data
    [
      'description',
      'templateFeatureId',
      'schemaXml',
      'defaultContentApprovalWorkflowId',
      'defaultDisplayFormUrl',
      'defaultEditFormUrl',
      'emailAlias',
      'sendToLocationName',
      'sendToLocationUrl',
      'validationFormula',
      'validationMessage' 
    ].forEach(o => {
      const value: any = (args.options as any)[o];
      if (value) {
        telemetryProps[o] = (typeof value !== 'undefined').toString();
      }
    });
    
    // add boolean values
    SpoListAddCommand.booleanOptions.forEach(o => {
      const value: any = (args.options as any)[o];
      if (value) {
        telemetryProps[o] = (value === 'true').toString();
      }
    });

    // add properties with non-identifiable data
    [
      'baseTemplate',
      'direction',
      'draftVersionVisibility',
      'listExperienceOptions',
      'majorVersionLimit',
      'majorWithMinorVersionsLimit',
      'readSecurity',
      'writeSecurity'
    ].forEach(o => {
      const value: any = (args.options as any)[o];
      if (value) {
        telemetryProps[o] = value.toString();
      }
    });*/

    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let listRestUrl: string = (args.options.listId ? 
        `${args.options.webUrl}/_api/web/lists/(guid'${args.options.listId}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${args.options.listTitle}')`);
    let contentTypeId: string = '';
    let listRootFolder: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.debug) {
          cmd.log(`Getting content types for list...`);
        }
        
        const requestOptions: any = {
          url: `${listRestUrl}/contenttypes`,
          method: 'GET',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);

      })
      .then((response: any): request.RequestPromise | Promise<void> => {

        if (this.debug) {
          cmd.log('content type lookup response...');
          cmd.log(response);
        }

        if (response.value.length && response.value.length > 0) {

          // Use first defined content type in list if no content type specified in arguments
          if (!args.options.contentType) {
            contentTypeId = response.value[0].Id.StringValue
            if (this.debug) {
              cmd.log(`using list's first content type id: ... ${contentTypeId}`);
            }
          } else {
            contentTypeId = response.value.filter( (ct: any) => { 
              ct['Id'].StringValue == args.options.contentType
              || ct['Name'] == args.options.contentType
            }) || { Id: { StringValue: '' }}.Id.StringValue
          }
        }

        if (this.debug) {
          cmd.log(`using content type id: ... ${contentTypeId}`);
        }
        
        if (args.options.folder) {
          
          if (this.debug) {
            cmd.log('setting up folder lookup response ...');
          }

          const requestOptions = { 
            url: `${listRestUrl}/rootFolder`,
            method: 'GET',
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'accept': 'application/json;odata=nometadata'
            }),
            json: true
          }

          if (this.debug) {
            cmd.log('Executing web request for list\'s root folder...');
            cmd.log(requestOptions);
            cmd.log('');
          }
          return request.get(requestOptions).then(rootFolderResponse => {

            if (this.debug) {
              cmd.log('list root folder lookup response...');
              cmd.log(rootFolderResponse);
            }

            listRootFolder = rootFolderResponse["ServerRelativeUrl"];
    
    
            return this.ensureFolder(args, siteAccessToken, cmd, rootFolderResponse, args.options.folder || "")
              .then(ensureFolderResponse => {

                if (this.debug) {
                  cmd.log('ensure folder response...');
                  cmd.log(ensureFolderResponse);
                }
    
              })
          }) 

  
        } else {
          return Promise.resolve();
        }

      })
      .then((response: any): request.RequestPromise => {

        if (this.verbose) {
          cmd.log(`Creating item in list ${args.options.listId || args.options.listTitle} on site at ${args.options.webUrl}...`);
        }

        const requestBody: any = { formValues: this.mapRequestBody(args.options) };

        if (args.options.folder) {
          requestBody["listItemCreateInfo"] = {"FolderPath": {"DecodedUrl":`${(listRootFolder + '/' + args.options.folder).replace(/\/\//gi, '/')}`} }
        }

        const requestOptions: any = {
          url: `${listRestUrl}/AddValidateUpdateItemUsingPath()`,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: requestBody,
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((response: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(response);
          cmd.log('');
        }

        if (response["data"]) {
          cmd.log(<ListItemInstance>response["data"])
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>',
        description: 'URL of the site where the item should be added'
      },
      {
        option: '-l, --listId <list guid>',
        description: 'GUID of the list where the item should be added'
      },
      {
        option: '--listTitle <listTitle>',
        description: 'Title of the list where the item should be added'
      },
      {
        option: '-c, --contentType <contentType>',
        description: 'The name or the ID of the content type to associate with the new item'
      },
      {
        option: '-f, --folder <folder>',
        description: 'The list-relative URL of the folder where the item should be created'
      },
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
        'contentType',
        'folder'
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
        return `Required parameters listId or listTitle missing`;
      }

      if (args.options.listId && args.options.listTitle) {
        return `Only specify one of listId or listTitle parameters`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To add an item to a list, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Add an item to a list with title ${chalk.grey('Demo Item')} and content type id ${chalk.grey('0x0101')}
    to list with title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LISTITEM_ADD} --title 'Demo Item' --contentType 0x0101 --listTitle 'Demo List' --webUrl https://contoso.sharepoint.com/sites/project-x

    More information:

    SP Client List Item Class Members information
      https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.listitem_members.aspx
   `);
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = [];
    const excludeOptions: string[] = [
      'listTitle',
      'listId',
      'webUrl',
      'contentType',
      'folder',
      'debug',
      'verbose'
  ]

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) == -1) {
        requestBody.push({FieldName: key, FieldValue: (<any>options)[key]});
      }
    });

    /*requestBody["__metadata"] = {
      type: `SP.Data.${(options.listTitle || "").replace(/\s/gi, "_x0020_")}ListItem`
    }*/
    /*if (options.writeSecurity) {
      requestBody.WriteSecurity = options.writeSecurity;
    }*/

    return requestBody;
  }

  private ensureFolder(args: any, siteAccessToken: string, cmd: any, rootFolder: any, folderToEnsure: string) {
    let rootFolderPath = rootFolder['ServerRelativeUrl']
    let childFolderNames = folderToEnsure.split('/');

    let checkFoldersPromise: (request.RequestPromise | Promise<void>)[] = [];
    let checkedFolders: string[] = [];
    let createFolders: any[] = [];

    for (let folderIndex = 0; folderIndex < childFolderNames.length; folderIndex++)
    {

      let folderName = childFolderNames[folderIndex];
      let parentFolders = folderIndex > 0 ? childFolderNames.slice(0, folderIndex) : []
      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${(rootFolderPath + '/' + parentFolders.join('/') + '/' + folderName).replace(/\/\//gi, "/")}')`,
        method: 'GET',
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${siteAccessToken}`,
          'accept': 'application/json;odata=nometadata'
        }),
        json: true
      };
      if (this.debug) {
        cmd.log('Setting up promise with web request to check folder...');
        cmd.log(requestOptions);
        cmd.log('');
      }
      checkFoldersPromise.push(new Promise((resolve, reject) => {
        request.get(requestOptions).then((response) => {

            if (this.debug) {
              cmd.log(`Folder ${folderName} found with response...`);
              cmd.log(response);
            }

            checkedFolders.push(folderName);
            resolve();

          }).catch( () => {

            createFolders.push({
              folderName: folderName,
              parentFolder: `${rootFolderPath}${(parentFolders.length > 0 ? '/' : '')}${parentFolders.join('/')}`
            });
            resolve();
          })

      }));
    }

    return Promise.all(checkFoldersPromise).then(() => {

      if (this.debug) {
        cmd.log(`Folders found:`);
        cmd.log(checkedFolders);
        cmd.log(`Folders to create:`);
        cmd.log(createFolders);
        cmd.log('');
      }

      let createFolderPromises: any[] = [];

      for (let i = 0; i < createFolders.length; i++) {

        // Below path is used by the modern UI
        const createFolderOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1='${(createFolders[i].parentFolder)}'&@a2='${(createFolders[i].folderName)}'`,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        }
  
        if (this.debug) {
          cmd.log('Setting up promise to create folder request...');
          cmd.log(createFolderOptions);
          cmd.log('');
        }
  
        createFolderPromises.push(request.post(createFolderOptions));

      }

      let counter = 0;
      return new Promise((resolve, reject) => {

        let recurse = () => {

          if (this.debug) {
            cmd.log(`Executing create folder promise ${counter}`);
          }
          createFolderPromises[counter].then((response: any) => {
            if (this.debug) {
              cmd.log(`Create folder promise ${counter} executed with response:`);
              cmd.log(response);
            }
            counter++;
            if (counter < createFolderPromises.length) {
              recurse();
            } else {
              resolve()
            }
          }).catch((response: any) => {
            if (this.debug) {
              cmd.log(`Error executing create folder promise ${counter}:`);
              cmd.log(response);
              cmd.log('');
            }
            counter++;
            if (counter < createFolderPromises.length) {
              recurse();
            } else {
              resolve();
            }
          });
  
        }

        if (createFolderPromises.length > 0) {
          recurse();
        } else {
          resolve();
        }
        

      })

    });

  }

}

module.exports = new SpoListItemAddCommand();