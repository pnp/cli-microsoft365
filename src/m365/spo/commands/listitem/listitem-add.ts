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
import { FolderExtensions } from '../../FolderExtensions';

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

interface FieldValue {
  ErrorMessage: string;
  FieldName: string;
  FieldValue: any;
  HasException: boolean;
  ItemId: number;
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
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.contentType = typeof args.options.contentType !== 'undefined';
    telemetryProps.folder = typeof args.options.folder !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);
    let contentTypeName: string = '';
    let targetFolderServerRelativeUrl: string = '';
    const folderExtensions: FolderExtensions = new FolderExtensions(cmd, this.debug);

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

    request
      .get(requestOptions)
      .then((response: any): Promise<void> => {
        if (args.options.contentType) {
          const foundContentType = response.value.filter((ct: any) => {
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

        if (args.options.folder) {
          if (this.debug) {
            cmd.log('setting up folder lookup response ...');
          }

          const requestOptions: any = {
            url: `${listRestUrl}/rootFolder`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            json: true
          }

          return request
            .get<any>(requestOptions)
            .then(rootFolderResponse => {
              targetFolderServerRelativeUrl = Utils.getServerRelativePath(rootFolderResponse["ServerRelativeUrl"], args.options.folder as string);

              return folderExtensions.ensureFolder(args.options.webUrl, targetFolderServerRelativeUrl);
            });
        }
        else {
          return Promise.resolve();
        }
      })
      .then((): Promise<any> => {
        if (this.verbose) {
          cmd.log(`Creating item in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }

        const requestBody: any = {
          formValues: this.mapRequestBody(args.options)
        };

        if (args.options.folder) {
          requestBody.listItemCreateInfo = {
            FolderPath: {
              DecodedUrl: targetFolderServerRelativeUrl
            }
          };
        }

        if (args.options.contentType && contentTypeName !== '') {
          if (this.debug) {
            cmd.log(`Specifying content type name [${contentTypeName}] in request body`);
          }

          requestBody.formValues.push({
            FieldName: 'ContentType',
            FieldValue: contentTypeName
          });
        }

        const requestOptions: any = {
          url: `${listRestUrl}/AddValidateUpdateItemUsingPath()`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          body: requestBody,
          json: true
        };

        return request.post(requestOptions);
      })
      .then((response: any): Promise<any> => {
        // Response is from /AddValidateUpdateItemUsingPath POST call, perform get on added item to get all field values
        const fieldValues: FieldValue[] = response.value;
        const idField = fieldValues.filter((thisField, index, values) => {
          return (thisField.FieldName == "Id");
        });

        if (this.debug) {
          cmd.log(`field values returned:`)
          cmd.log(fieldValues)
          cmd.log(`Id returned by AddValidateUpdateItemUsingPath: ${idField}`);
        }

        if (idField.length === 0) {
          return Promise.reject(`Item didn't add successfully`)
        }

        const requestOptions: any = {
          url: `${listRestUrl}/items(${idField[0].FieldValue})`,
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
        description: 'URL of the site where the item should be added'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list where the item should be added. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list where the item should be added. Specify listId or listTitle but not both'
      },
      {
        option: '-c, --contentType [contentType]',
        description: 'The name or the ID of the content type to associate with the new item'
      },
      {
        option: '-f, --folder [folder]',
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
  
    Add an item with Title ${chalk.grey('Demo Item')} and content type name ${chalk.grey('Item')} to list with
    title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --contentType Item --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"

    Add an item with Title ${chalk.grey('Demo Multi Managed Metadata Field')} and
    a single-select metadata field named ${chalk.grey('SingleMetadataField')} to list with
    title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Managed Metadata Field" --SingleMetadataField "TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;"

    Add an item with Title ${chalk.grey('Demo Multi Managed Metadata Field')} and a multi-select
    metadata field named ${chalk.grey('MultiMetadataField')} to list with title ${chalk.grey('Demo List')}
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Managed Metadata Field" --MultiMetadataField "TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;"
  
    Add an item with Title ${chalk.grey('Demo Single Person Field')} and a single-select people
    field named ${chalk.grey('SinglePeopleField')} to list with title ${chalk.grey('Demo List')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Person Field" --SinglePeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'}]"
      
    Add an item with Title ${chalk.grey('Demo Multi Person Field')} and a multi-select people
    field named ${chalk.grey('MultiPeopleField')} to list with title ${chalk.grey('Demo List')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Person Field" --MultiPeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'},{'Key':'i:0#.f|membership|adamb@conotoso.com'}]"
    
    Add an item with Title ${chalk.grey('Demo Hyperlink Field')} and a hyperlink field named
    ${chalk.grey('CustomHyperlink')} to list with title ${chalk.grey('Demo List')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Hyperlink Field" --CustomHyperlink "https://www.bing.com, Bing"
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
      'verbose',
      'output'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody.push({ FieldName: key, FieldValue: `${(<any>options)[key]}` });
      }
    });
    return requestBody;
  }
}

module.exports = new SpoListItemAddCommand();