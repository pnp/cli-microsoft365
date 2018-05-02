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
import { FolderBaseCommand } from './folder-base';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  recycle?: boolean;
  confirm?: boolean;
}

class SpoFolderRemoveCommand extends FolderBaseCommand {
  public get name(): string {
    return commands.FOLDER_REMOVE;
  }

  public get description(): string {
    return 'Deletes a folder form a site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.recycle = (!(!args.options.recycle)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    const removeFolder: () => void = (): void => {
      if (this.debug) {
        cmd.log(`Retrieving access token for ${resource}...`);
      }

      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise => {
          siteAccessToken = accessToken;

          if (this.verbose) {
            cmd.log(`Removing folder in site at ${args.options.webUrl}...`);
          }

          let webRelativeUrl = this.getWebRelativeUrlFromWebUrl(args.options.webUrl);
          let serverRelativeUrl: string = `${webRelativeUrl}${this.formatRelativeUrl(args.options.sourceUrl)}`;
          let requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')`;
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
        .then((resp: any): void => {
          if(this.verbose){
            cmd.log('DONE');  
          }
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeFolder();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the folder ${args.options.sourceUrl} located in site ${args.options.webUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeFolder();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folder will be deleted'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>',
        description: 'Site-relative URL of the target folder'
      },
      {
        option: '--recycle',
        description: 'Recycles the folder instead of actually deleting it'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the folder'
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

      if (!args.options.sourceUrl) {
        return 'Required parameter sourceUrl missing';
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
  
    To remove a folder, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    The remove folder command will remove folder only if it is empty. This is how the SharePoint REST and client.svc APIs work.
        
  Examples:

    Removes a folder with site-relative URL ${chalk.grey('/Shared Documents/My Folder')} located in
    site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FOLDER_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --sourceUrl '/Shared Documents/My Folder'

    Moves a folder with site-relative URL ${chalk.grey('/Shared Documents/My Folder')} located in
    site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} to the recycle bin of the site
      ${chalk.grey(config.delimiter)} ${commands.FOLDER_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --sourceUrl '/Shared Documents/My Folder' --recycle
    `)
  }
}

module.exports = new SpoFolderRemoveCommand();