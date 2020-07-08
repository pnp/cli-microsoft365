import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { FolderProperties } from './FolderProperties';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  parentFolderUrl: string;
  name: string;
}

class SpoFolderAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ADD;
  }

  public get description(): string {
    return 'Creates a folder within a parent folder';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Adding folder to site ${args.options.webUrl}...`);
    }

    const parentFolderServerRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.parentFolderUrl);
    const serverRelativeUrl: string = `${parentFolderServerRelativeUrl}/${args.options.name}`;
    const requestUrl: string = `${args.options.webUrl}/_api/web/folders`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata',
      },
      body: {
        'ServerRelativeUrl': serverRelativeUrl
      },
      json: true
    };

    request
      .post<FolderProperties>(requestOptions)
      .then((folder: FolderProperties): void => {
        cmd.log(folder);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folder will be created'
      },
      {
        option: '-p, --parentFolderUrl <parentFolderUrl>',
        description: 'Site-relative URL of the parent folder'
      },
      {
        option: '-n, --name <name>',
        description: 'Name of the new folder to be created'
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

      if (!args.options.parentFolderUrl) {
        return 'Required parameter parentFolderUrl missing';
      }

      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Creates folder in a parent folder with site relative url ${chalk.grey('/Shared Documents')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.FOLDER_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --parentFolderUrl '/Shared Documents' --name 'My Folder Name'
    `);
  }
}

module.exports = new SpoFolderAddCommand();