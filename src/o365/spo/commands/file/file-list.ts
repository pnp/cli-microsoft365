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
import { Auth } from '../../../../Auth';
import { FilePropertiesCollection } from './FilePropertiesCollection';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folder: string;
}

class SpoFileListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_LIST;
  }

  public get description(): string {
    return 'Lists all available files in the specified folder and site';
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
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving all files in folder ${args.options.folder} at site ${args.options.webUrl}...`);
        }

        let requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files`;

        if (args.options.output !== 'json') {
          requestUrl += '?$select=UniqueId,Name,ServerRelativeUrl';
        }

        const requestOptions: any = {
          url: requestUrl,
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
      .then((fileProperties: FilePropertiesCollection): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(fileProperties);
          cmd.log('');
        }

        cmd.log(fileProperties.value);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folder from which to retrieve files is located'
      },
      {
        option: '-f, --folder <folder>',
        description: 'The server- or site-relative URL of the folder from which to retrieve files'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required option webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.folder) {
        return 'Required option folder missing';
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
  
    To get all files, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Return all files from folder ${chalk.grey('Shared Documents')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_LIST} --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents'
      `);
  }
}

module.exports = new SpoFileListCommand();