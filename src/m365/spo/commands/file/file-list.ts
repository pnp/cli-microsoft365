import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderProperties } from '../folder/FolderProperties.js';
import { FileProperties } from './FileProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  folderUrl: z.string().alias('f'),
  fields: z.string().optional(),
  filter: z.string().optional(),
  recursive: z.boolean().optional().alias('r')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all files in folder '${args.options.folderUrl}' at site '${args.options.webUrl}'${args.options.recursive ? ' (recursive)' : ''}...`);
    }

    try {
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

      await logger.log(allFiles);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFiles(folderUrl: string, fieldProperties: FieldProperties, args: CommandArgs, logger: Logger, skip: number = 0): Promise<FileProperties[]> {
    if (this.verbose) {
      const page = Math.ceil(skip / SpoFileListCommand.pageSize) + 1;
      await logger.logToStderr(`Retrieving files in folder '${folderUrl}'${page > 1 ? ', page ' + page : ''}...`);
    }

    const allFiles: FileProperties[] = [];
    const serverRelativePath: string = urlUtil.getServerRelativePath(args.options.webUrl, folderUrl);
    const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/Files`;
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
      url: `${requestUrl}?${queryParams.join('&')}`,
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
      await logger.logToStderr(`Retrieving folders in folder '${folderUrl}'${page > 1 ? ', page ' + page : ''}...`);
    }

    const allFolders: string[] = [];
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, folderUrl);
    const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/Folders`;

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$skip=${skip}&$top=${SpoFileListCommand.pageSize}&$select=ServerRelativeUrl`,
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

export default new SpoFileListCommand();