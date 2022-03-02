import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileFolderCollection } from '../folder/FileFolderCollection';
import { FileProperties } from './FileProperties';
import { FilePropertiesCollection } from './FilePropertiesCollection';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folder: string;
  recursive?: boolean;
}

class SpoFileListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_LIST;
  }

  public get description(): string {
    return 'Lists all available files in the specified folder and site';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all files in folder ${args.options.folder} at site ${args.options.webUrl}...`);
    }

    this.getFiles(args.options.folder, args).then((files: FilePropertiesCollection): void => {
      logger.log(files.value);
      cb();
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  // Gets files from a folder recursively.
  private getFiles(folderUrl: string, args: CommandArgs, files: FilePropertiesCollection = { value: [] }): Promise<FilePropertiesCollection> {
    // If --recursive option is specified, retrieve both Files and Folder details, otherwise only Files.
    const expandParameters: string = args.options.recursive ? 'Files,Folders' : 'Files';
    let requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderUrl)}')?$expand=${expandParameters}`;
    if (args.options.output !== 'json') {
      requestUrl += '&$select=Files/UniqueId,Files/Name,Files/ServerRelativeUrl';
    }
    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<FileFolderCollection>(requestOptions)
      .then((filesAndFoldersResult: FileFolderCollection) => {
        filesAndFoldersResult.Files.forEach((file: FileProperties) => files.value.push(file));
        // If the request is --recursive, call this method for other folders.
        if (args.options.recursive &&
          filesAndFoldersResult.Folders !== undefined &&
          filesAndFoldersResult.Folders.length !== 0) {
          return Promise.all(filesAndFoldersResult.Folders.map((folder: { ServerRelativeUrl: string; }) => this.getFiles(folder.ServerRelativeUrl, args, files)));
        }
        else {
          return;
        }
      }).then(() => files);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --folder <folder>'
      },
      {
        option: '-r, --recursive'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoFileListCommand();