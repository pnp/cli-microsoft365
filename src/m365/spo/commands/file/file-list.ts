import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderProperties } from '../folder/FolderProperties';
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
  private static readonly thresholdLimit = 5000;
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
        option: '--fields [fields]'
      },
      {
        option: '--filter [filter]'
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
      // If --recursive option is specified, retrieve both Files and Folder details, otherwise only Files.
      const subfolderFiles: FileProperties[] = [];
      if (args.options.recursive) {
        const subFolders = await this.getSubfoldersPages(args.options.folder, args, 0);

        for (const folder of subFolders) {
          const subfolderFilesForFolder: FilePropertiesCollection = await this.getFiles(folder, args, 0);
          subfolderFilesForFolder.value.forEach((file: FileProperties) => subfolderFiles.push(file));
        }
      }

      const files = await this.getFiles(args.options.folder, args, 0);
      logger.log([...files.value, ...subfolderFiles]);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  // Gets files from a folder recursively.
  private async getFiles(folderUrl: string, args: CommandArgs, index: number, files: FilePropertiesCollection = { value: [] }): Promise<FilePropertiesCollection> {
    const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderUrl)}')/Files`;

    const fieldsProperties = this.formatSelectProperties(args.options.fields, args.options.output);

    const queryParams = [`$skip=${index}`, `$top=${SpoFileListCommand.thresholdLimit}`];

    if (fieldsProperties.expandProperties.length > 0) {
      queryParams.push(`$expand=${fieldsProperties.expandProperties.join(',')}`);
    }

    if (fieldsProperties.selectProperties.length > 0) {
      queryParams.push(`$select=${fieldsProperties.selectProperties.join(',')}`);
    }

    if (args.options.filter) {
      queryParams.push(`$filter=${args.options.filter}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?${queryParams.join('&')}`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const filesAndFoldersResult = await request.get<{ value: FileProperties[] }>(requestOptions);
    filesAndFoldersResult.value.forEach((file: FileProperties) => files.value.push(file));

    if (filesAndFoldersResult.value.length === SpoFileListCommand.thresholdLimit) {
      await this.getFiles(folderUrl, args, index + SpoFileListCommand.thresholdLimit, files);
    }

    return files;
  }

  private async getSubfoldersPages(folderUrl: string, args: CommandArgs, index: number, folders: string[] = []): Promise<string[]> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderUrl)}')/Folders?$skip=${index}&$top=${SpoFileListCommand.thresholdLimit}`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const resp = await request.get<{ value: FolderProperties[] }>(requestOptions);
    if (resp.value.length > 0) {
      if (resp.value.length === SpoFileListCommand.thresholdLimit) {
        await this.getSubfoldersPages(folderUrl, args, index + SpoFileListCommand.thresholdLimit, folders);
      }
      for (const folder of resp.value) {
        folders.push(folder.ServerRelativeUrl);
        await this.getSubfoldersPages(folder.ServerRelativeUrl, args, 0, folders);
      }
    }

    return folders;
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