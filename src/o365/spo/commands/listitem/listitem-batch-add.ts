import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
//import { ListItemInstance } from './ListItemInstance';
import { FolderExtensions } from '../../FolderExtensions';

import * as path from 'path';


const vorpal: Vorpal = require('../../../../vorpal-init');
const csv = require('fast-csv');
const fs = require('fs');


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  contentType?: string;
  folder?: string;
  path: string;
}

// interface FieldValue {
//   ErrorMessage: string;
//   FieldName: string;
//   FieldValue: any;
//   HasException: boolean;
//   ItemId: number;
// }

class SpoListItemAddCommand extends SpoCommand {

  public allowUnknownOptions(): boolean | undefined {
    return false;
  }

  public get name(): string {
    return commands.LISTITEM_BATCH_ADD;
  }

  public get description(): string {
    return 'Creates a list item in the specified list for each row in the specified .csv file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.contentType = typeof args.options.contentType !== 'undefined';
    telemetryProps.folder = typeof args.options.folder !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let lineNumber: number = 0;
    let contentTypeName: string | null = null;
    let listRestUrl: string | null = null;
    let batchSize: number = 250;
    let UuidDenerator = this.generateUUID;
    let recordsToAdd = new Array();




    let targetFolderServerRelativeUrl: string = ``;
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = Utils.getSafeFileName(path.basename(fullPath));
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    listRestUrl = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);



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
      .then((): any => {
        if (this.verbose) {
          cmd.log(`Creating items in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }
        // get the file
        let fileStream = fs.createReadStream(fileName);
        let lmapRequestBody = this.mapRequestBody;
        fileStream.on('error', function () {
          cb(`error accessing file ${fileName}`);
        });
        let csvStream: any = fileStream.pipe(csv.parse({ headers: true }));
        //start the batch
        let changeSetId = this.generateUUID();
        let endpoint = `${listRestUrl}/AddValidateUpdateItemUsingPath()`;
        let linesInBatch = 0;
        csvStream
          .on("data", function (row: any) {
            lineNumber++; linesInBatch++;
            const requestBody: any = {
              formValues: lmapRequestBody(row)
            };
            if (args.options.folder) {
              requestBody.listItemCreateInfo = {
                FolderPath: {
                  DecodedUrl: targetFolderServerRelativeUrl
                }
              };
            }
            if (args.options.contentType && contentTypeName !== '') {
              requestBody.formValues.push({
                FieldName: 'ContentType',
                FieldValue: contentTypeName
              });
            }
            // row is ready
            recordsToAdd.push('--changeset_' + changeSetId);
            recordsToAdd.push('Content-Type: application/http');
            recordsToAdd.push('Content-Transfer-Encoding: binary');
            recordsToAdd.push('');
            recordsToAdd.push('POST ' + endpoint + ' HTTP/1.1');
            recordsToAdd.push('Content-Type: application/json;odata=verbose');
            recordsToAdd.push('');
            recordsToAdd.push(`${JSON.stringify(requestBody)}`);
            recordsToAdd.push('');
            cmd.log(`line #= ${lineNumber}; lines in batch : ${linesInBatch} batchsue ${batchSize}`)
            if (linesInBatch >= batchSize) {
              /***
               *  Send the batch
               * 
               */
              cmd.log(`sending batch`)
              recordsToAdd.push('--changeset_' + changeSetId + '--');
              csvStream.pause();
              var batchBody = recordsToAdd.join('\u000d\u000a');
              let batchContents = new Array();
              let batchId = UuidDenerator();
              batchContents.push('--batch_' + batchId);
              batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
              batchContents.push('Content-Length: ' + batchBody.length);
              batchContents.push('Content-Transfer-Encoding: binary');
              batchContents.push('');
              batchContents.push(batchBody);
              batchContents.push('');
              const updateOptions: any = {
                url: `${args.options.webUrl}/_api/$batch`,
                headers: {
                  //'X-RequestDigest': formDigestValue,
                  'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
                },
                body: batchContents.join('\r\n')
              }
              cmd.log(updateOptions)
              request.post(updateOptions)
                .catch((e) => {
                  cmd.log(`error`);
                  cb(e);
                })
                .then((results) => {
                  cmd.log(`results`);
                  cmd.log(results);
                })

                .finally(() => {
                  cmd.log(`vatrcgh done`);
                  recordsToAdd.splice(0,recordsToAdd.length);//clear the buffer
                  changeSetId = UuidDenerator();
                  csvStream.resume();
                  linesInBatch = 0;
                })
             

            }
          })
          .on("end", function () {
            if (linesInBatch > 0){
              cmd.log(`sending final batch`)
              recordsToAdd.push('--changeset_' + changeSetId + '--');
              var batchBody = recordsToAdd.join('\u000d\u000a');
              let batchContents = new Array();
              let batchId = UuidDenerator();
              batchContents.push('--batch_' + batchId);
              batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
              batchContents.push('Content-Length: ' + batchBody.length);
              batchContents.push('Content-Transfer-Encoding: binary');
              batchContents.push('');
              batchContents.push(batchBody);
              batchContents.push('');
              const updateOptions: any = {
                url: `${args.options.webUrl}/_api/$batch`,
                headers: {
                  //'X-RequestDigest': formDigestValue,
                  'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
                },
                body: batchContents.join('\r\n')
              }
              cmd.log(updateOptions)
              request.post(updateOptions)
                .catch((e) => {
                  cmd.log(`error`);
                  cb(e);
                })
                .then((results) => {
                  cmd.log(`results`);
                  cmd.log(results);
                })

                .finally(() => {
                  cmd.log(`Processed ${lineNumber} Rows`)
                  cmd.log(`vatrcgh done`);
                  cb();
                })
            }else{
              cmd.log(`Processed ${lineNumber} Rows`)

              cb();
            }
            
            
          })
          .on("error", function (error: any) {
            cmd.log(error)
          });

      })
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the item should be added'
      },
      {
        option: '-p, --path <path>',
        description: 'the path of the csv file with records to be added to the SharePoint list'
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


  private generateUUID() {
    var d = new Date().getTime();
    var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = (d + Math.random() * 16) % 16 | 0;
      d = Math.floor(d / 16);
      return (c == 'x' ? r : (r & 0x7 | 0x8)).toString(16);
    });
    return uuid;
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

      if (!args.options.path) {
        return `Specify path`;
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

  private mapRequestBody(row: any): any {
    const requestBody: any = [];
    Object.keys(row).forEach(async key => {
      requestBody.push({ FieldName: key, FieldValue: (<any>row)[key] });
    });
    return requestBody;
  }
}

module.exports = new SpoListItemAddCommand();