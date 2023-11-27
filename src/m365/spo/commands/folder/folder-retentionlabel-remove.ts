import * as url from 'url';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderProperties } from './FolderProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl?: string;
  folderId?: string;
  force?: boolean;
}

class SpoFolderRetentionLabelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_RETENTIONLABEL_REMOVE;
  }

  public get description(): string {
    return 'Clear the retention label from a folder';
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
        folderUrl: typeof args.options.folderUrl !== 'undefined',
        folderId: typeof args.options.folderId !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--folderUrl [folderUrl]'
      },
      {
        option: '-i, --folderId [folderId]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.folderId &&
          !validation.isValidGuid(args.options.folderId as string)) {
          return `${args.options.folderId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['folderUrl', 'folderId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeFolderRetentionLabel(logger, args);
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the retentionlabel from folder ${args.options.folderId || args.options.folderUrl} located in site ${args.options.webUrl}?` });

      if (result) {
        await this.removeFolderRetentionLabel(logger, args);
      }
    }
  }

  private async removeFolderRetentionLabel(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const folderProperties = await this.getFolderProperties(logger, args);

      if (folderProperties.ListItemAllFields) {
        const parsedUrl = url.parse(args.options.webUrl);
        const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
        const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, folderProperties.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);

        await spo.removeRetentionLabelFromListItems(args.options.webUrl, listAbsoluteUrl, [parseInt(folderProperties.ListItemAllFields.Id)], logger, args.options.verbose);
      }
      else {
        const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, folderProperties.ServerRelativeUrl);

        await spo.removeDefaultRetentionLabelFromList(args.options.webUrl, listAbsoluteUrl, logger, args.options.verbose);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderProperties(logger: Logger, args: CommandArgs): Promise<FolderProperties> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list and item information for folder '${args.options.folderId || args.options.folderUrl}' in site at ${args.options.webUrl}...`);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.folderId) {
      requestUrl += `GetFolderById('${args.options.folderId}')`;
    }
    else {
      const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl!);
      requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<FolderProperties>(requestOptions);
  }
}

export default new SpoFolderRetentionLabelRemoveCommand();