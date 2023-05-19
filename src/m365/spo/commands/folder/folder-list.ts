import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
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

interface FieldProperties {
  selectProperties: string[];
  expandProperties: string[];
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
        option: '-f, --fields [fields]'
      },
      {
        option: '--filter [filter]'
      },
      {
        option: '-r, --recursive [recursive]'
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
      logger.logToStderr(`Retrieving all folders in folder '${args.options.parentFolderUrl}' at site '${args.options.webUrl}'${args.options.recursive ? ' (recursive)' : ''}...`);
    }

    try {
      const fieldProperties = this.formatSelectProperties(args.options.fields);
      const allFiles = await this.getFolders(args.options.parentFolderUrl, fieldProperties, args, logger);

      // Clean ListItemAllFields.ID property from the output if included
      // Reason: It causes a casing conflict with 'Id' when parsing JSON in PowerShell
      if (fieldProperties.selectProperties.some(p => p.toLowerCase().indexOf('listitemallfields') > -1)) {
        allFiles.filter(folder => (folder.ListItemAllFields as any)?.ID !== undefined).forEach(folder => delete (folder.ListItemAllFields as any)['ID']);
      }

      logger.log(allFiles);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolders(parentFolderUrl: string, fieldProperties: FieldProperties, args: CommandArgs, logger: Logger, skip: number = 0): Promise<FolderProperties[]> {
    if (this.verbose) {
      const page = Math.ceil(skip / SpoFolderListCommand.pageSize) + 1;
      logger.logToStderr(`Retrieving folders in folder '${parentFolderUrl}'${page > 1 ? ', page ' + page : ''}...`);
    }

    const allFolders: FolderProperties[] = [];
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, parentFolderUrl);
    const requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl(@url)/Folders?@url='${formatting.encodeQueryParameter(serverRelativeUrl)}'`;
    const queryParams = [`$skip=${skip}`, `$top=${SpoFolderListCommand.pageSize}`];

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

    const response = await request.get<{ value: FolderProperties[] }>(requestOptions);

    for (const folder of response.value) {
      allFolders.push(folder);

      if (args.options.recursive) {
        const subFolders = await this.getFolders(folder.ServerRelativeUrl, fieldProperties, args, logger);
        subFolders.forEach(subFolder => allFolders.push(subFolder));
      }
    }

    if (response.value.length === SpoFolderListCommand.pageSize) {
      const folders = await this.getFolders(parentFolderUrl, fieldProperties, args, logger, skip + SpoFolderListCommand.pageSize);
      folders.forEach(folder => allFolders.push(folder));
    }

    return allFolders;
  }

  private formatSelectProperties(fields: string | undefined): FieldProperties {
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