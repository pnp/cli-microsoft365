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
import { v4 } from 'uuid';
import { createReadStream } from 'fs';
import requestPromise = require('request-promise-native');


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
  batchSize: number;
}
interface FieldNames {
  value: { InternalName: string }[]
}
class SpoListItemAddCommand extends SpoCommand {

  public allowUnknownOptions(): boolean | undefined {
    return;
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
  public static parseResults(response: string, cmd: CommandInstance, cb: (err?: any) => void): void {
    cmd.log(typeof response)
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

  public static mapRequestBody(row: any, csvHeaders: Array<string>): any {
    const requestBody: any = [];
    Object.keys(row).forEach(async key => {
      requestBody.push({ FieldName: csvHeaders[parseInt(key)], FieldValue: (<any>row)[key] });
    });
    return requestBody;
  }
  public static sendABatch(batchCounter: number, rowsInBatch: number, changeSetId: string, recordsToAdd: string, webUrl: string, verbose: boolean, cmd: CommandInstance): Promise<string> {

    const batchContents = new Array();
    const batchId = v4();
    batchContents.push('--batch_' + batchId);
    if (verbose) {
      cmd.log(`Sending batch #${batchCounter} with ${rowsInBatch} items`);
    }
    batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
    batchContents.push('Content-Length: ' + recordsToAdd.length);
    batchContents.push('Content-Transfer-Encoding: binary');
    batchContents.push('');
    batchContents.push(recordsToAdd);

    batchContents.push('--batch_' + batchId + '--');

    const updateOptions: requestPromise.OptionsWithUrl = {
      url: `${webUrl}/_api/$batch`,
      headers: {
        'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
      },
      body: batchContents.join('\r\n')
    };
    cmd.log(updateOptions)
    return request.post(updateOptions);
  }
  public static async   validateContentType(contentTypeName: string | undefined, listRestUrl: string, webUrl: string, verbose: boolean, cmd: CommandInstance): Promise<any> {
    if (contentTypeName == undefined) {
      return (Promise.resolve());
    }
    if (verbose) {
      cmd.log(`Getting content types for list...`);
    }
    const ctRequestOptions: requestPromise.OptionsWithUrl = {
      url: `${listRestUrl}/contenttypes?$select=Name,Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    return request
      .get(ctRequestOptions)
      .then((response: any): Promise<void> => {

        const foundContentType = response.value.filter((ct: any) => {
          const contentTypeMatch: boolean = ct.Id.StringValue === contentTypeName || ct.Name === contentTypeName;
          if (verbose) {
            cmd.log(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
          }
          return contentTypeMatch;
        });
        if (verbose) {
          cmd.log('content type filter output...');
          cmd.log(foundContentType);
        }
        if (foundContentType.length !== 1) {
          return Promise.reject(`Specified content type '${contentTypeName}' doesn't exist on the target list`);
        } else {
          return (Promise.resolve())
        }


      })
  }
  public static async   getFolderUrl(folderName: string | undefined, listRestUrl: string, webUrl: string, verbose: boolean, debug:boolean, cmd: CommandInstance): Promise<any> {
    if (folderName == undefined) {
      cmd.log(listRestUrl)
      const listRequestOptions: requestPromise.OptionsWithUrl = {
        url: listRestUrl+"/RootFolder",
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };
      return request
        .get(listRequestOptions)
        .then((response: any): Promise<string> => {
          cmd.log(response)
          return Promise.resolve(response.ServerRelativeUrl);
        })

    }
    else{
 
      const requestOptions: requestPromise.OptionsWithUrl = {
        url: `${listRestUrl}/rootFolder`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      }
      return request
        .get<any>(requestOptions)
        .then(async (rootFolderResponse) => {

          const targetFolderServerRelativeUrl = Utils.getServerRelativePath(rootFolderResponse["ServerRelativeUrl"], folderName);
          const folderExtensions: FolderExtensions = new FolderExtensions(cmd, debug);
          await folderExtensions.ensureFolder(webUrl, targetFolderServerRelativeUrl);
          return targetFolderServerRelativeUrl;

        });
    }
  }
  public static async   getCaseSensitiveWebUrl( webUrl: string, cmd: CommandInstance): Promise<any> {

      const WebRequestOptions: requestPromise.OptionsWithUrl = {
        url: webUrl+"/_api/web",
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };
      return request
        .get(WebRequestOptions)
        .then((response: any): Promise<string> => {
          cmd.log(response)
          return Promise.resolve(response.Url);
        })

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
//    let targetFolderServerRelativeUrl: string = ``;
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = Utils.getSafeFileName(path.basename(fullPath));
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const batchSize: number = args.options.batchSize || 10;
    listRestUrl = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    
    SpoListItemAddCommand.validateContentType(args.options.contentType, listRestUrl, args.options.webUrl, this.verbose, cmd)
      .catch((ctError) => {
        cb(ctError)
        cmd.log("error on ct")
      })
      .then((): Promise<string | void> => {
        return SpoListItemAddCommand.getCaseSensitiveWebUrl(args.options.webUrl,cmd);
         
       })
      .then((caseCorrectedWebUrl:string|void): Promise<string | void> => {
        listRestUrl = (args.options.listId ?
          `${caseCorrectedWebUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
          : `${caseCorrectedWebUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);
    
       return SpoListItemAddCommand.getFolderUrl(args.options.folder,listRestUrl as string,caseCorrectedWebUrl as string,this.verbose,this.debug,cmd);
        
      })
      .then((folderServerRelativeUrl: string | void): any => {
        
        if (this.verbose) {
          cmd.log(`Creating items in list ${folderServerRelativeUrl}`);
        }
        //start the batch -- each batch will get assigned its own id
        let changeSetId = v4();
        const endpoint = `${listRestUrl}/AddValidateUpdateItemUsingPath()`;
        // get the csv  file passed in from the cmd line
        let fileStream = createReadStream(fileName);
        let csvStream: any = csv.parseStream(fileStream, { headers: false });
        let verboseMode = this.verbose;
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
                const fetchFieldsRequest: requestPromise.OptionsWithUrl = {
                  url: `${listRestUrl}/fields?$select=InternalName&$filter=ReadOnlyField eq false`,
                  json: true,
                  headers: {
                    'Accept': `application/json;odata=nometadata`
                  },
                }
                request.get<FieldNames>(fetchFieldsRequest)
                  .then((realFields: FieldNames) => {

                    for (let header of csvHeaders) {
                      let fieldFound = false;
                      for (let spField of realFields.value) {
                        if (header === spField.InternalName) {
                          fieldFound = true;
                          break;
                        }
                      }
                      if (!fieldFound) {
                        cmd.log(`Field ${header} was not found on the SharePoint list.  Valid fields follow`)
                        cmd.log(realFields)
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
                  formValues: SpoListItemAddCommand.mapRequestBody(row, csvHeaders)
                };
                if (args.options.folder) {
                  requestBody.listItemCreateInfo = {
                    FolderPath: {
                      DecodedUrl: folderServerRelativeUrl
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
                recordsToAdd += '--changeset_' + changeSetId + '\r\n' +
                  'Content-Type: application/http' + '\r\n' +
                  'Content-Transfer-Encoding: binary' + '\r\n' +
                  '\r\n' +
                  'POST ' + endpoint + ' HTTP/1.1' + '\r\n' +
                  'Content-Type: application/json;odata=verbose' + '\r\n' +
                  'Accept: application/json;odata=verbose' + '\r\n' +
                  '\r\n' +
                  `${JSON.stringify(requestBody)}` + '\r\n' +
                  '\r\n';

                /***  Send the batch if the buffer is getting full   **/
                if (rowsInBatch >= batchSize || recordsToAdd.length >= maxBytesInBatch) {

                  recordsToAdd += '--changeset_' + changeSetId + '--' + '\r\n';
                  ++batchCounter;
                  if (verboseMode) {
                    cmd.log(`Sending batch #${batchCounter} with ${rowsInBatch} items`)
                  }

                  SpoListItemAddCommand.sendABatch(batchCounter, rowsInBatch, changeSetId, recordsToAdd, args.options.webUrl, verboseMode, cmd)
                    .catch((e) => {
                      cb(e);
                    })
                    .then((response: string | void) => {

                      SpoListItemAddCommand.parseResults(response as string, cmd, cb)
                      recordsToAdd = ``;
                      rowsInBatch = 0;
                      changeSetId = v4();
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
              ++batchCounter;
              recordsToAdd += '--changeset_' + changeSetId + '--' + '\r\n';
              if (verboseMode) {
                cmd.log(`Sending final batch #${batchCounter} with ${rowsInBatch} items`)
              }

              SpoListItemAddCommand.sendABatch(batchCounter, rowsInBatch, changeSetId, recordsToAdd, args.options.webUrl, verboseMode, cmd)
                .catch((e) => {
                  cb(e);
                })
                .then((response: string | void) => {
                  SpoListItemAddCommand.parseResults(response as string, cmd, cb)
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
        description: 'The maximum number of records to send to SharePoint in a batch (default is 10)'
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
        args.options.batchSize > 1000) {
        return `batchsize ${args.options.batchSize} exceeds the 1000 item limit`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
    The first row of the csv file contains column headers. The column headers must match the internal name of the field.
    If you send a file with an invalid fieldname the command will display an error that includes all the valid field names:
      Field Assignee was not found on the SharePoint list.  Valid fields follow
      InternalName
      ------------
      ContentType
      Title
      Attachments
      Order
      FileLeafRef
      MetaInfo

    The rows in the csv must contain column values based on the type of the column:
      Text: The text to be added to the column 
      Number: the number to be added to the column
      Single-Select Metadata: the metadata name folowed by the pipe (|) charachter, followed by the metadata Id followed by a semicolon (i.e. TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377; )
      Multie-Select Metadata: the metadata name folowed by the pipe (|) charachter, followed by the metadata Id followed by a semicolon. This is repeated for each term. (i.e. ermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75; )
      Single-Select Person: {'Key':'i:0#.f|membership|--UPN--'}  where --UPN-- is the UPN of the person to be added
      Multi-Select Person: [{'Key':'i:0#.f|membership|--UPN1--'},{'Key':'i:0#.f|membership|--UPN2--'}]  where --UPN1-- and --UPN2-- are the UPNs of the persons to be added.
      Hyperlink: the  url of the hyperlink followd bt the text to be displayed for the hyperlink. This must be enclosed in quotes. (i.e. "https://www.bing.com, Bing")


  
  Examples:
  
    Add an item with Title ${chalk.grey('Demo Item')} and content type name ${chalk.grey('Item')} to list with
    title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_BATCH_ADD} --listTitle "Test" --webUrl https://contoso.sharepoint.com/sites/project-x --path .\test.csv

   `);
  }


}

module.exports = new SpoListItemAddCommand();