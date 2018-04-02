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
import Utils from '../../../../Utils';
import Auth from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  url?: string;
  recycle?: boolean;
  confirm?: boolean;
}

class SpoFileRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    // it may have identifiable information
    // telemetryProps.url = (!(!args.options.url)).toString();
        
    telemetryProps.recycle = (!(!args.options.recycle)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    const removeFile: () => void = (): void => {
      if (this.debug) {
        cmd.log(`Retrieving access token for ${resource}...`);
      }

      // concatenate trailing '/' if not provided
      // so if the provided url is for the root site, the substr bellow will get the right value
      let serverRelativeSiteUrl: string = args.options.webUrl;
      if (serverRelativeSiteUrl[serverRelativeSiteUrl.length-1] !== '/') {
        serverRelativeSiteUrl = `${serverRelativeSiteUrl}/`;
      } 
      serverRelativeSiteUrl = serverRelativeSiteUrl.substr(serverRelativeSiteUrl.indexOf('/', 8));
      
      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise => {
          siteAccessToken = accessToken;

          if (this.verbose) {
            cmd.log(`Removing file in site at ${args.options.webUrl}...`);
          }

          let requestUrl: string = '';

          if (args.options.id) {
            requestUrl = `${args.options.webUrl}/_api/web/GetFileById(guid'${encodeURIComponent(args.options.id as string)}')`;
          }
          else {
            let fileUrl = args.options.url as string;
            if (!fileUrl.startsWith(serverRelativeSiteUrl)) {
              fileUrl = `${serverRelativeSiteUrl}${fileUrl}`
            }
            requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`;
          }

          if (args.options.recycle) {
            requestUrl += `/recycle()`;
          }

          const requestOptions: any = {
            url: requestUrl,
            method: 'POST',
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-HTTP-Method': 'DELETE',
              'If-Match': '*',
              'accept': 'application/json;odata=nometadata'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.post(requestOptions);
        })
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeFile();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.recycle? "recycle": "remove"} the file ${args.options.id || args.options.url} located at site ${args.options.webUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeFile();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the file to remove is located'
      },
      {
        option: '-i, --id [id]',
        description: 'The ID of the file to remove. Specify either id or url but not both'
      },
      {
        option: '-u, --url [url]',
        description: 'The server or site-relative url of the file to remove. Specify either id or url but not both'
      },      
      {
        option: '--recycle',
        description: 'Recycle the file instead of actually deleting it'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the file'
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

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id as string)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.id && args.options.url) {
        return 'Specify id or url, but not both';
      }

      if (!args.options.id && !args.options.url) {
        return 'Specify id or url';
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
  
    To remove a file, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Remove the file with ID ${chalk.grey('0cd891ef-afce-4e55-b836-fce03286cccf')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --id 0cd891ef-afce-4e55-b836-fce03286cccf

    Remove the file with Site Relative Url ${chalk.grey('SharedDocuments/Test.docx')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --url SharedDocuments/Test.docx
      `);
  }
}

module.exports = new SpoFileRemoveCommand();