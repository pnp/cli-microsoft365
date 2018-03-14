import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo } from '../../spo';
import Utils from '../../../../Utils';
import Auth from '../../../../Auth';
import fs = require('fs');
import * as path from 'path';
import { FileProperties } from './FileProperties';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  url: string;
  id: string;
  asString: boolean;
  asListItem: boolean;
  asFile: boolean;
  fileName: string;
  path: string;
}

class FileGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_GET;
  }

  public get description(): string {
    return 'Get information about the specified file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.url = (!(!args.options.url)).toString();
    telemetryProps.asString = args.options.asString || false;
    telemetryProps.asListItem = args.options.asListItem || false;
    telemetryProps.asFile = args.options.asFile || false;
    telemetryProps.fileName = (!(!args.options.fileName)).toString();
    telemetryProps.path = (!(!args.options.path)).toString();

    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(`Retrieving file from site ${args.options.webUrl}...`);
        }

        let requestUrl: string;
        let options: string;

        if (args.options.id) {
          requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${encodeURIComponent(args.options.id)}')`;
        }
        else {
          requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(args.options.url)}')`;
        }

        if (args.options.asListItem) {
          options = '?$expand=ListItemAllFields';
        }
        else if (!args.options.asListItem && !args.options.asFile && !args.options.asString) {
          options = '';
        }
        else {
          options = '/$value';
        }

        const requestOptions: any = {
          url: requestUrl + options,
          method: 'GET',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          encoding: null, // Set encoding to null, otherwise binary data will be encoded to utf8 and binary data is corrupt 
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((file: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(file);
          cmd.log('');
        }
        
        if (args.options.asString) {
          cmd.log(file);
        }
        else if (args.options.asListItem) {
          let fileProperties: FileProperties = JSON.parse(JSON.stringify(file));
          cmd.log(fileProperties.ListItemAllFields)
        }
        else if(args.options.asFile) {
          this.writeFile(file, args.options.path, args.options.fileName)
          cmd.log(`File ${args.options.fileName} saved to path ${args.options.path}`);
        }
        else {
          let fileProperties: FileProperties = JSON.parse(JSON.stringify(file));
          cmd.log(fileProperties);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private writeFile(fileContent: string, filePath: string, fileName: string): void {
    let fullPath: string = path.join(filePath, fileName);

    fs.writeFileSync(fullPath, fileContent);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>',
        description: 'The URL of the site where the folder from which to retrieve files is located'
      },
      {
        option: '-u, --url [url]',
        description: 'server- or site-relative URL of the file. Specify either url or id but not both'
      },
      {
        option: '-i, --id [id]',
        description: 'file ID. Specify either url or id but not both'
      },
      {
        option: '--asString',
        description: 'retrieve the contents of the specified file as string'
      },
      {
        option: '--asListItem',
        description: 'retrieve the underlying list item'
      },
      {
        option: '--asFile',
        description: 'save the file to the path specified in the path option'
      },
      {
        option:'-f, --fileName [fileName]',
        description: 'the name of the file including extension. Must be specified when the --asFile option is used'
      },
      {
        option: '-p, --path [path]',
        description: 'path where to save the file. Must be specified when the --asFile option is used'
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

      if (args.options.id) {
        if (!Utils.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
      }

      if (args.options.id && args.options.url) {
        return 'Specify id or url, but not both';
      }

      if (!args.options.id && !args.options.url) {
        return 'Specify id or url, one is required';
      }

      if (args.options.asFile && (!args.options.path || !args.options.fileName)) {
        return 'The path and fileName should be specified when the --asFile option is used';
      }

      if (args.options.path && !fs.existsSync(args.options.path)) {
        return 'Specified path does not exits';
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
  
    To get a file, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Return file properties for file with id ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'

    Return file as string for file with id ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asString

    Return list item properties for file with id ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asListItem   

    Save file at path ${chalk.grey('/Users/user/documents')} with filename ${chalk.grey('SavedAsTest1.docx')} for file with id ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asFile --path /Users/user/documents --fileName SavedAsTest1.docx
    
    Return file properties for file with site relative url ${chalk.grey('/sites/project-x/documents/Test1.docx')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx'

    Return file as string for file with site relative url ${chalk.grey('/sites/project-x/documents/Test1.docx')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asString

    Return list item properties for file with site relative url ${chalk.grey('/sites/project-x/documents/Test1.docx')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asListItem   

    Save file at path ${chalk.grey('/Users/user/documents')} with filename ${chalk.grey('SavedAsTest1.docx')} for file with site relative url ${chalk.grey('/sites/project-x/documents/Test1.docx')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_GET} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asFile --path /Users/user/documents --fileName SavedAsTest1.docx
      `);
  }
}

module.exports = new FileGetCommand();