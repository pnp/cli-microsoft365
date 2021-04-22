import * as url from 'url';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetFilename: string;
  force?: boolean;
  
}

class SpoFileRenameCommand extends SpoCommand {
  //private dots?: string;

  public get name(): string {
    return commands.FILE_RENAME;
  }

  public get description(): string {
    return 'Renames a file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.deleteIfAlreadyExists = args.options.deleteIfAlreadyExists || false;
    telemetryProps.allowSchemaMismatch = args.options.allowSchemaMismatch || false;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const webUrl = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
    const targetFileUrl: string = this.createTargetFileUrl(args.options.sourceUrl,args.options.targetFilename);
    const targetFileSiteUrl: string = webUrl.replace(tenantUrl,'');
   
     // Check if the source file exists.
    // Called on purpose, we explicitly check if user specified file
    // in the sourceUrl option.
    // use MoveTo to rename a file

    this
      .fileExists(tenantUrl, webUrl, args.options.sourceUrl)
      .then((): Promise<void> => {
        if (args.options.force) {
          // try delete target file, if force flag is set
         
          return this.recycleFile(webUrl, targetFileSiteUrl, targetFileUrl, logger);
        }

        return Promise.resolve();
      })
      .then((): Promise<any> => {
        // all preconditions met, let's rename the file        
        let targetFileUrlReplaced: string  = "";
        let sourceUrlReplaced: string = "";
        let targetFileSiteUrlReplaced = "";
        
        if (targetFileUrl.lastIndexOf('/') !== targetFileUrl.length - 1) {
          targetFileUrlReplaced = `${targetFileUrl}/`;
        }
        if (args.options.sourceUrl.lastIndexOf('/') !== args.options.sourceUrl.length - 1) {
           sourceUrlReplaced = `${args.options.sourceUrl}/`;
       }
       if (targetFileSiteUrl.lastIndexOf('/') !== targetFileSiteUrl.length - 1) {
        targetFileSiteUrlReplaced = `${targetFileSiteUrl}/`;
       }
       const endpointUrl: string = `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(targetFileSiteUrlReplaced)}${encodeURIComponent(sourceUrlReplaced)}')/moveto(newurl='${encodeURIComponent(targetFileSiteUrlReplaced)}${encodeURIComponent(targetFileUrlReplaced)}',flags=1)`;
       const requestOptions: any = {
        url: endpointUrl,
        method: 'PUT',
        headers: {
          'X-HTTP-Method': 'PUT',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };
       return request.post(requestOptions);
      })
      .then((jobInfo: any): Promise<any> => {
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
         
        });
      })
      .then((): void => {
        if (this.verbose) {
          logger.logToStderr('DONE');
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));

   
    
  }
     
  

   private fileExists(tenantUrl: string, webUrl: string, sourceUrl: string): Promise<void> {
    const webServerRelativeUrl: string = webUrl.replace(tenantUrl, '');
    const fileServerRelativeUrl: string = `${webServerRelativeUrl}${sourceUrl}`;    
    const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelativeUrl)}')/`;
    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    

    return request.get(requestOptions);
  }

  private createTargetFileUrl(sourceUrl: string, targetFileName: string): string {

    const targetFileUrl: string = sourceUrl.substring(0,sourceUrl.lastIndexOf("/"))+ "/"+targetFileName;
    return targetFileUrl;

  }

  private recycleFile(webUrl: string, targetUrl: string,targetFileUrl: string , logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const targetFolderAbsoluteUrl: string = this.urlCombine(webUrl, targetUrl);
      //https://cgikislay.sharepoint.com/sites/contosoportal
      ///Shared%20Documents/contoso_report1.pptx

      // since the target WebFullUrl is unknown we can use getRequestDigest
      // to get it from target folder absolute url.
      // Similar approach used here Microsoft.SharePoint.Client.Web.WebUrlFromFolderUrlDirect
      this.getRequestDigest(targetFolderAbsoluteUrl)
        .then((contextResponse: ContextInfo): void => {
          if (this.debug) {
            logger.logToStderr(`contextResponse.WebFullUrl: ${contextResponse.WebFullUrl}`);
          }

          if (targetUrl.charAt(0) !== '/') {
            targetUrl = `/${targetUrl}`;
          }
          if (targetUrl.lastIndexOf('/') !== targetUrl.length - 1) {
            targetUrl = `${targetUrl}/`;
          }
          if (targetFileUrl.lastIndexOf('/') !== targetFileUrl.length - 1) {
            targetFileUrl = `${targetFileUrl}/`;
          }

          const requestUrl: string = `${contextResponse.WebFullUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(`${targetUrl}${targetFileUrl}`)}')/recycle()`;
          const requestOptions: any = {
            url: requestUrl,
            method: 'POST',
            headers: {
              'X-HTTP-Method': 'DELETE',
              'If-Match': '*',
              'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          request.post(requestOptions)
            .then((): void => {
              resolve();
            })
            .catch((err: any): any => {
              if (err.statusCode === 404) {
                // file does not exist so can proceed
                return resolve();
              }

              if (this.debug) {
                logger.logToStderr(`recycleFile error...`);
                logger.logToStderr(err);
              }

              reject(err);
            });
        }, (e: any) => reject(e));
    });
  }
 
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the file is located'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>',
        description: 'Site-relative URL of the file to move'
      },
      {
        option: '-t, --targetFilename <targetFilename>',
        description: 'New Filename for the file to be renamed'
      },
      {
        option: '--force',
        description: 'If a file already exists at the targetUrl, it will be moved to the recycle bin. If omitted, the move operation will be canceled if the file already exists at the targetUrl location'
      },
      
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoFileRenameCommand();
