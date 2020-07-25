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
import { CommandInstance } from '../../../../cli';

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
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoFolderAddCommand();