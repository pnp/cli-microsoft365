import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { FolderProperties } from './FolderProperties';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
}

class SpoFolderGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified folder';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving folder from site ${args.options.webUrl}...`);
    }

    const serverRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
    const requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<FolderProperties>(requestOptions)
      .then((folder: FolderProperties): void => {
        cmd.log(folder);

        cb();
      }, (err: any): void => {
        if (err.statusCode && err.statusCode === 500) {
          cb(new CommandError('Please check the folder URL. Folder might not exist on the specified URL'));
          return;
        }

        this.handleRejectedODataJsonPromise(err, cmd, cb);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folder is located'
      },
      {
        option: '-f, --folderUrl <folderUrl>',
        description: 'Site-relative URL of the folder'
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

module.exports = new SpoFolderGetCommand();