import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileProperties } from './FileProperties';
import { FilePropertiesCollection } from './FilePropertiesCollection';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folder: string;
  recursive?: boolean;
  fields?: string;
  filter?: string;
}

class SpoFileListCommand extends SpoCommand {
  private static readonly tresholdLimit = 5000;
  public get name(): string {
    return commands.FILE_LIST;
  }

  public get description(): string {
    return 'Lists all available files in the specified folder and site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        recursive: args.options.recursive,
        fields: typeof args.options.fields !== 'undefined',
        filter: typeof args.options.filter !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --folder <folder>'
      },
      {
        option: '-r, --recursive'
      },
      {
        option: '-f, --fields [fields]'
      },
      {
        option: '-l, --filter [filter]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all files in folder ${args.options.folder} at site ${args.options.webUrl}...`);
    }

    try {
      const folderProperties = await this.getItemCount(args.options.folder, args);
      const files = await this.getFiles(args.options.folder, args, folderProperties.folders, folderProperties.items, 0);
      logger.log(files.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  // Gets files from a folder recursively.
  private async getFiles(folderUrl: string, args: CommandArgs, subFolders: any[], items: number, index: number, files: FilePropertiesCollection = { value: [] }): Promise<FilePropertiesCollection> {
    // If --recursive option is specified, retrieve both Files and Folder details, otherwise only Files.
    //const expandParameters: string = args.options.recursive ? 'Files,Folders' : 'Files';
    const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderUrl)}')/Files`;

    const fieldsProperties = this.formatSelectProperties(args.options.fields, args.options.output);

    const options = [`$skip=${index}`];

    if (!args.options.recursive && items > SpoFileListCommand.tresholdLimit) {
      options.push(`$top=${SpoFileListCommand.tresholdLimit}`);
    }
    else {
      options.push(`$top=${items}`);
    }

    options.push();

    if (fieldsProperties.expandProperties.length > 0) {
      options.push(`$expand=${fieldsProperties.expandProperties.join(',')}`);
    }

    if (fieldsProperties.selectProperties.length > 0) {
      options.push(`$select=${fieldsProperties.selectProperties.join(',')}`);
    }

    if (args.options.filter) {
      options.push(`$filter=${args.options.filter}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?${options.join('&')}`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: FileProperties[] }>(requestOptions)
      .then(async (filesAndFoldersResult: { value: FileProperties[] }) => {
        filesAndFoldersResult.value.forEach((file: FileProperties) => files.value.push(file));
        // If the request is --recursive, call this method for other folders.
        if (args.options.recursive &&
          subFolders.length !== 0) {
          return Promise.all(subFolders.map(async (folder: string) => {
            const folderProperties = await this.getItemCount(folder, args);
            this.getFiles(folder, args, folderProperties.folders, folderProperties.items, 0, files);
          }));
        }
        else if (items > SpoFileListCommand.tresholdLimit && (items - index > SpoFileListCommand.tresholdLimit)) {
          return await this.getFiles(folderUrl, args, [], items, index + SpoFileListCommand.tresholdLimit, files);
        }
        else {
          return;
        }
      }).then(() => files);
  }

  private async getItemCount(folderUrl: string, args: CommandArgs): Promise<{ items: number, folders: any[] }> {
    let expandProperties = 'Properties';
    if (args.options.recursive) {
      expandProperties += ',Folders';
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${formatting.encodeQueryParameter(folderUrl)}')?$expand=${expandProperties}&$select=Properties/vti_x005f_folderitemcount,Properties/vti_x005f_foldersubfolderitemcount`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response: any = await request.get(requestOptions);

    const urls: any[] = [];
    if (response.Folders && response.Folders.length > 0) {
      response.Folders.forEach((folder: any) => {
        if (folder.ServerRelativeUrl) {
          urls.push(folder.ServerRelativeUrl);
        }
      });
    }

    return { items: response.Properties.vti_x005f_folderitemcount - response.Properties.vti_x005f_foldersubfolderitemcount, folders: urls };
  }

  private formatSelectProperties(fields: string | undefined, output: string | undefined): { selectProperties: string[], expandProperties: string[] } {

    let selectProperties: any[] = [];
    const expandProperties: any[] = [];

    if (output !== 'json') {
      selectProperties = ['UniqueId', 'Name', 'ServerRelativeUrl'];
    }

    if (fields) {
      fields.split(',').forEach((field) => {
        const subparts = field.trim().split('/');
        if (subparts.length > 1) {
          expandProperties.push(subparts[0]);
        }
        selectProperties.push(field.trim());
      });
    }

    return {
      selectProperties: [...new Set(selectProperties)],
      expandProperties: [...new Set(expandProperties)]
    };
  }

}

module.exports = new SpoFileListCommand();