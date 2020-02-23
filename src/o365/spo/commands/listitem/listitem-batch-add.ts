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
import { Transform } from 'stream';
const vorpal: Vorpal = require('../../../../vorpal-init');
const csv = require('@fast-csv/parse');
import { createReadStream } from 'fs';

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
  batchSize:number;
}

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
    let maxBytesInBatch: number = 1000000; // max is  1048576
    let rowsInBatch: number = 0;
    let batchCounter = 0;
      let recordsToAdd = "";
    let csvHeaders: Array<string>;
    const generateUUID = function () {
      var d = new Date().getTime();
      var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = (d + Math.random() * 16) % 16 | 0;
        d = Math.floor(d / 16);
        return (c == 'x' ? r : (r & 0x7 | 0x8)).toString(16);
      });
      return uuid;
    }
    const parseResults = (response: any, cmd: CommandInstance, cb: (err?: any) => void): void => {
      let responseLines: Array<string> = response.toString().split('\n');
      // read each line until you find JSON... 
      for (let responseLine of responseLines) {
        try {
          //check for error 

          if (responseLine.startsWith("HTTP/1.1 5")) { //any 500 errors (like timeout), just stop
            cmd.log("An HTTP 5xx error was returned from SharePoint. Please retry with a lower --batchsize ")
            cb(responseLine);
          }
          // parse the JSON response...
          var tryParseJson = JSON.parse(responseLine);
          for (let result of tryParseJson.d.AddValidateUpdateItemUsingPath.results) {
            if (result.HasException) {
              cmd.log(result)
            }
          }
        } catch (e) {
        }
      }
    }

    const mapRequestBody = (row: any, csvHeaders: Array<string>): any => {
      const requestBody: any = [];
      Object.keys(row).forEach(async key => {
        requestBody.push({ FieldName: csvHeaders[parseInt(key)], FieldValue: (<any>row)[key] });
      });
      return requestBody;
    }
    let targetFolderServerRelativeUrl: string = ``;
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = Utils.getSafeFileName(path.basename(fullPath));
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const batchSize: number=args.options.batchSize|| 10;
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
        //start the batch -- each batch will get assigned its own id
        let changeSetId = generateUUID();
        let endpoint = `${listRestUrl}/AddValidateUpdateItemUsingPath()`;
        // get the csv  file passed in from the cmd line
        let fileStream = createReadStream(fileName);
        let csvStream: any = csv.parseStream(fileStream, { headers: false })
        csvStream
          .pipe(new Transform({ //https://github.com/C2FO/fast-csv/issues/328 Need to transform if  we are batching asynch
            objectMode: true,
            write(row: any, encoding: string, callback: (error?: (Error | null)) => void): void {
              if (lineNumber === 0) {
                /***
                 * Process csv Headers (fast csv headers doens not work if using transform)
                 */
                csvHeaders = row;
                // fetch the valid field names from the list. If you pass a bad field name to AddValidateUpdateItemUsingPath it returns xml not JSON
                const fetchFieldsRequest: any = {
                  url: `${listRestUrl}/fields?$select=InternalName&$filter=ReadOnlyField eq false`,
                  headers: {
                    'Accept': `application/json;odata=nometadata`
                  },
                }
                request.get(fetchFieldsRequest)
                  .then((realFields) => {
                    const spFields = JSON.parse(realFields as string).value
                    for (let header of csvHeaders) {
                      let fieldFound = false;
                      for (let spField of spFields) {
                        if (header === spField.InternalName) {
                          fieldFound = true;
                          break;
                        }
                      }
                      if (!fieldFound) {
                        cmd.log(`Field ${header} was not found on the SharePoint list.  Valid fields follow`)
                        cmd.log(spFields)
                        cb(`Error-- field ${header} was not found on the SharePoint list`)
                      }
                    }
                    lineNumber++
                    this.push(row);
                    callback();
                  })
                  .catch((error) => {
                    cb(error)
                  })

              }
              else {
                /***
                * Process csv Data
                */
                lineNumber++
                rowsInBatch++;
                const requestBody: any = {
                  formValues: mapRequestBody(row, csvHeaders)
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
                recordsToAdd += '--changeset_' + changeSetId + '\u000d\u000a' +
                  'Content-Type: application/http' + '\u000d\u000a' +
                  'Content-Transfer-Encoding: binary' + '\u000d\u000a' +
                  '\u000d\u000a' +
                  'POST ' + endpoint + ' HTTP/1.1' + '\u000d\u000a' +
                  'Content-Type: application/json;odata=verbose' + '\u000d\u000a' +
                  'Accept: application/json;odata=verbose' + '\u000d\u000a' +
                  '\u000d\u000a' +
                  `${JSON.stringify(requestBody)}` + '\u000d\u000a' +
                  '\u000d\u000a';

                /***  Send the batch if the buffer is getting full   **/
                if (rowsInBatch >= batchSize || recordsToAdd.length >= maxBytesInBatch) {

                  recordsToAdd += '--changeset_' + changeSetId + '--' + '\u000d\u000a';
                  let batchContents = new Array();
                  let batchId = generateUUID();
                  batchContents.push('--batch_' + batchId);
                  cmd.log(`Sending batch #${++batchCounter} with ${rowsInBatch} items`)
                  batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
                  batchContents.push('Content-Length: ' + recordsToAdd.length);
                  batchContents.push('Content-Transfer-Encoding: binary');
                  batchContents.push('');
                  batchContents.push(recordsToAdd);
                  batchContents.push('');
                  const updateOptions: any = {
                    url: `${args.options.webUrl}/_api/$batch`,
                    headers: {
                      'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
                    },
                    body: batchContents.join('\r\n')
                  }
                  request.post(updateOptions)
                    .catch((e) => {
                      cb(e);
                    })
                    .then((response) => {
                      parseResults(response, cmd, cb)
                      recordsToAdd = ``;
                      rowsInBatch = 0;
                      changeSetId = generateUUID();
                      this.push(row);
                      callback();
                    })
                }
                else {
                  this.push(row);
                  callback();
                }

              }
            },
          }))
          .on("data", function () { })// dont delete this ,  or on end wont fire
          .on("end", function () {

            if (recordsToAdd.length > 0) {


              let batchContents = new Array();
              let batchId = generateUUID();
              batchContents.push('--batch_' + batchId);
              cmd.log(`Sending final batch #${++batchCounter} with ${rowsInBatch} items`)

              batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
              batchContents.push('Content-Length: ' + recordsToAdd.length);
              batchContents.push('Content-Transfer-Encoding: binary');
              batchContents.push('');
              batchContents.push(recordsToAdd);
              batchContents.push('');
              const updateOptions: any = {
                url: `${args.options.webUrl}/_api/$batch`,
                headers: {
                  'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
                },
                body: batchContents.join('\r\n')
              }
              request.post(updateOptions)
                .catch((e) => {
                  cb(e);
                })
                .then((response) => {
                  parseResults(response, cmd, cb)
                })
                .finally(() => {
                  cmd.log(`Processed ${lineNumber} Rows`)
                  cb();
                })
            } else {
              cmd.log(`Processed ${lineNumber} Rows`)
              cb();
            }
          })
          .on("error", function (error: any) {
            cb(error)
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
      {
        option: '-b, --batchSize [batchSize]',
        description: 'The maximum number of records to sent to SharePoint in a batch (default is 400)'
      },
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
      if (args.options.batchSize &&
        args.options.batchSize > 1000){
        return `batchsize ${args.options.batchSize} exceeds the 1000 item limit`;
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


}

module.exports = new SpoListItemAddCommand();