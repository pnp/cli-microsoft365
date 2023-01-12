import { AxiosRequestConfig } from 'axios';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderProperties } from './FolderProperties';
import { Options as SpoListItemRetentionLabelRemoveCommandOptions } from '../listitem/listitem-retentionlabel-remove';
import * as SpoListItemRetentionLabelRemoveCommand from '../listitem/listitem-retentionlabel-remove';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl?: string;
  folderId?: string;
  confirm?: boolean;
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
        confirm: !!args.options.confirm
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
        option: '--confirm'
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
    if (args.options.confirm) {
      await this.removeFolderRetentionLabel(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the retentionlabel from folder ${args.options.folderId || args.options.folderUrl} located in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await this.removeFolderRetentionLabel(logger, args);
      }
    }
  }

  private async removeFolderRetentionLabel(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing retention label from folder ${args.options.folderId || args.options.folderUrl} in site at ${args.options.webUrl}...`);
    }
    try {
      const folderProperties = await this.getFolderProperties(args);
      const options: SpoListItemRetentionLabelRemoveCommandOptions = {
        webUrl: args.options.webUrl,
        listId: folderProperties.listId,
        listItemId: folderProperties.id,
        confirm: true,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      const spoListItemRetentionLabelRemoveCommandOutput = await Cli.executeCommandWithOutput(SpoListItemRetentionLabelRemoveCommand as Command, { options: { ...options, _: [] } });
      if (this.verbose) {
        logger.logToStderr(spoListItemRetentionLabelRemoveCommandOutput.stderr);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderProperties(args: CommandArgs): Promise<{ id: string, listId: string }> {
    const requestOptions: AxiosRequestConfig = {
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.folderId) {
      requestUrl += `GetFolderById('${args.options.folderId}')`;
    }
    else {
      const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl!);
      requestUrl += `GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }

    requestOptions.url = `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ListItemAllFields/ParentList/Id,ListItemAllFields/Id`;
    const response = await request.get<FolderProperties>(requestOptions);

    return { id: response.ListItemAllFields.Id, listId: response.ListItemAllFields.ParentList.Id };
  }
}

module.exports = new SpoFolderRetentionLabelRemoveCommand();