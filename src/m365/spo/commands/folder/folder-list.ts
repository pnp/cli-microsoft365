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
}

class SpoFolderListCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_LIST;
  }

  public get description(): string {
    return 'Returns all folders under the specified parent folder';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving folders from site ${args.options.webUrl} parent folder ${args.options.parentFolderUrl}...`);
    }

    const serverRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.parentFolderUrl);
    const requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/folders`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<{ value: FolderProperties[] }>(requestOptions)
      .then((resp: { value: FolderProperties[] }): void => {
        if (args.options.output === 'json') {
          cmd.log(resp.value);
        }
        else {
          cmd.log(resp.value.map(f => {
            return {
              Name: f.Name,
              ServerRelativeUrl: f.ServerRelativeUrl
            }
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folders to list are located'
      },
      {
        option: '-p, --parentFolderUrl <parentFolderUrl>',
        description: 'Site-relative URL of the parent folder'
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

module.exports = new SpoFolderListCommand();