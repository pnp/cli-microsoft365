import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { FilePropertiesCollection } from './FilePropertiesCollection';
import { CommandInstance } from '../../../../cli';

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
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<FilePropertiesCollection>(requestOptions)
      .then((fileProperties: FilePropertiesCollection): void => {
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
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoFileListCommand();