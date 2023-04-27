import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderProperties } from '../folder/FolderProperties';
import { FileProperties } from './FileProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folder?: string;
  folderUrl?: string;
  recursive?: boolean;
  fields?: string;
  filter?: string;
}

interface FieldProperties {
  selectProperties: string[];
  expandProperties: string[];
}

class SpoFileListCommand extends SpoCommand {
  private static readonly pageSize = 5000;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        folder: typeof args.options.folder !== 'undefined',
        folderUrl: typeof args.options.folderUrl !== 'undefined',
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
        option: '-f, --folder [folder]'
      },
      {
        option: '-f, --folderUrl [folderUrl]'
      },
      {
        option: '--fields [fields]'
      },
      {
        option: '--filter [filter]'
      },
      {
        option: '-r, --recursive'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['folder', 'folderUrl'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all files in folder '${args.options.folder}' at site '${args.options.webUrl}'${args.options.recursive ? ' (recursive)' : ''}...`);
    }

    try {
      if (args.options.folder) {
        args.options.folderUrl = args.options.folder;

        this.warn(logger, `Option 'folder' is deprecated. Please use 'folderUrl' instead`);
      }

      const fieldProperties = this.formatSelectProperties(args.options.fields, args.options.output);
      const allFiles: FileProperties[] = [];
      const allFolders: string[] = args.options.recursive
        ? [...await this.getFolders(args.options.folderUrl!, args, logger), args.options.folderUrl!]
        : [args.options.folderUrl!];

      for (const folder of allFolders) {
        const files: FileProperties[] = await this.getFiles(folder, fieldProperties, args, logger);
        files.forEach((file: FileProperties) => allFiles.push(file));
      }

      // Clean ListItemAllFields.ID property from the output if included
      // Reason: It causes a casing conflict with 'Id' when parsing JSON in PowerShell
      if (fieldProperties.selectProperties.some(p => p.toLowerCase().indexOf('listitemallfields') > -1)) {
        allFiles.filter(file => file.ListItemAllFields?.ID !== undefined).forEach(file => delete file.ListItemAllFields['ID']);
      }

      logger.log(allFiles);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFiles(folderUrl: string, fieldProperties: FieldProperties, args: CommandArgs, logger: Logger, skip: number = 0): Promise<FileProperties[]> {
    if (this.verbose) {
      const page = Math.ceil(skip / SpoFileListCommand.pageSize) + 1;
      logger.logToStderr(`Retrieving files in folder '${folderUrl}'${page > 1 ? ', page ' + page : ''}...`);
    }

    const allFiles: FileProperties[] = [];
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, folderUrl);
    const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Files?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'`;
    const queryParams = [`$skip=${skip}`, `$top=${SpoFileListCommand.pageSize}`];

    if (fieldProperties.expandProperties.length > 0) {
      queryParams.push(`$expand=${fieldProperties.expandProperties.join(',')}`);
    }

    if (fieldProperties.selectProperties.length > 0) {
      queryParams.push(`$select=${fieldProperties.selectProperties.join(',')}`);
    }

    if (args.options.filter) {
      queryParams.push(`$filter=${args.options.filter}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}&${queryParams.join('&')}`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: FileProperties[] }>(requestOptions);
    response.value.forEach(file => allFiles.push(file));

    if (response.value.length === SpoFileListCommand.pageSize) {
      const files: FileProperties[] = await this.getFiles(folderUrl, fieldProperties, args, logger, skip + SpoFileListCommand.pageSize);
      files.forEach(file => allFiles.push(file));
    }

    return allFiles;
  }

  private async getFolders(folderUrl: string, args: CommandArgs, logger: Logger, skip: number = 0): Promise<string[]> {
    if (this.verbose) {
      const page = Math.ceil(skip / SpoFileListCommand.pageSize) + 1;
      logger.logToStderr(`Retrieving folders in folder '${folderUrl}'${page > 1 ? ', page ' + page : ''}...`);
    }

    const allFolders: string[] = [];
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, folderUrl);
    const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'`;

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}&$skip=${skip}&$top=${SpoFileListCommand.pageSize}&$select=ServerRelativeUrl`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: FolderProperties[] }>(requestOptions);

    for (const folder of response.value) {
      allFolders.push(folder.ServerRelativeUrl);
      const subfolders = await this.getFolders(folder.ServerRelativeUrl, args, logger);
      subfolders.forEach(folder => allFolders.push(folder));
    }

    if (response.value.length === SpoFileListCommand.pageSize) {
      const folders = await this.getFolders(folderUrl, args, logger, skip + SpoFileListCommand.pageSize);
      folders.forEach(folder => allFolders.push(folder));
    }

    return allFolders;
  }

  private formatSelectProperties(fields: string | undefined, output: string | undefined): FieldProperties {
    let selectProperties: any[] = [];
    const expandProperties: any[] = [];

    if (output === 'text' && !fields) {
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