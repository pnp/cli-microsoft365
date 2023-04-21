import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderProperties } from './FolderProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  parentFolderUrl: string;
  recursive?: boolean;
  fields?: string;
  filter?: string;
}

class SpoFolderListCommand extends SpoCommand {
  private static readonly pageSize = 5000;

  public get name(): string {
    return commands.FOLDER_LIST;
  }

  public get description(): string {
    return 'Returns all folders under the specified parent folder';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'ServerRelativeUrl'];
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
        recursive: !!args.options.recursive,
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
        option: '-p, --parentFolderUrl <parentFolderUrl>'
      },
      {
        option: '--recursive [recursive]'
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
      logger.logToStderr(`Retrieving folders from site ${args.options.webUrl} parent folder ${args.options.parentFolderUrl} ${args.options.recursive ? '(recursive)' : ''}...`);
    }

    try {
      // const folderProperties = await this.getItemCount(args.options.parentFolderUrl, args);

      // +1 since there is a hidden 'Forms' folder
      const resp = await this.getFolders(args.options.parentFolderUrl, args);
      logger.log(resp);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolders(parentFolderUrl: string, args: CommandArgs, skip: number = 0): Promise<FolderProperties[]> {
    let folders: FolderProperties[] = [];
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, parentFolderUrl);

    const fieldsProperties = this.formatSelectProperties(args.options.fields);

    const options = [`$skip=${skip}`];

    options.push(`$top=${SpoFolderListCommand.pageSize}`);

    if (fieldsProperties.expandProperties.length > 0) {
      options.push(`$expand=${fieldsProperties.expandProperties.join(',')}`);
    }

    if (fieldsProperties.selectProperties.length > 0) {
      options.push(`$select=${fieldsProperties.selectProperties.join(',')}`);
    }

    if (args.options.filter) {
      options.push(`$filter=${args.options.filter}`);
    }

    const resp = await odata.getAllItems<FolderProperties>(`${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/folders?${options.join('&')}`);
    if (resp.length > 0) {
      for (const folder of resp) {
        folders.push(folder);
        if (args.options.recursive) {
          const subfolders = await this.getFolders(folder.ServerRelativeUrl, args);
          folders = [...folders, ...subfolders];
        }
      }
    }

    if (resp.length === SpoFolderListCommand.pageSize) {
      const subfolders = await this.getFolders(parentFolderUrl, args, skip + SpoFolderListCommand.pageSize);
      folders = [...folders, ...subfolders];
    }

    return folders;
  }

  // private async getItemCount(folderUrl: string, args: CommandArgs): Promise<{ items: number }> {
  //   const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, folderUrl);
  //   const expandProperties = 'Properties';

  //   const requestOptions: CliRequestOptions = {
  //     url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${formatting.encodeQueryParameter(serverRelativeUrl)}')?$expand=${expandProperties}&$select=Properties/vti_x005f_folderitemcount,Properties/vti_x005f_foldersubfolderitemcount`,
  //     method: 'GET',
  //     headers: {
  //       'accept': 'application/json;odata=nometadata'
  //     },
  //     responseType: 'json'
  //   };

  //   const response: any = await request.get(requestOptions);

  //   return { items: response.Properties.vti_x005f_foldersubfolderitemcount };
  // }

  private formatSelectProperties(fields: string | undefined): { selectProperties: string[], expandProperties: string[] } {
    const selectProperties: any[] = [];
    const expandProperties: any[] = [];

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

module.exports = new SpoFolderListCommand();