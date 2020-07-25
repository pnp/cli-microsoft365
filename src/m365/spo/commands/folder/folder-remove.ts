import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  recycle?: boolean;
  confirm?: boolean;
}

class SpoFolderRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_REMOVE;
  }

  public get description(): string {
    return 'Deletes the specified folder';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.recycle = (!(!args.options.recycle)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeFolder: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Removing folder in site at ${args.options.webUrl}...`);
      }

      const serverRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
      let requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')`;
      if (args.options.recycle) {
        requestUrl += `/recycle()`;
      }
      const requestOptions: any = {
        url: requestUrl,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .post(requestOptions)
        .then((): void => {
          if (this.verbose) {
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
        message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the folder ${args.options.folderUrl} located in site ${args.options.webUrl}?`,
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
        description: 'The URL of the site where the folder to be deleted is located'
      },
      {
        option: '-f, --folderUrl <folderUrl>',
        description: 'Site-relative URL of the folder to delete'
      },
      {
        option: '--recycle',
        description: 'Recycles the folder instead of actually deleting it'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming deleting the folder'
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

module.exports = new SpoFolderRemoveCommand();