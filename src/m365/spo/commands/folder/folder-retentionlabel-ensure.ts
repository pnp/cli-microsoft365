import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListRetentionLabelEnsureCommand, { Options as SpoListRetentionLabelEnsureCommandOptions } from '../list/list-retentionlabel-ensure.js';
import spoListItemRetentionLabelEnsureCommand, { Options as SpoListItemRetentionLabelEnsureCommandOptions } from '../listitem/listitem-retentionlabel-ensure.js';
import { FolderProperties } from './FolderProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  folderUrl?: string;
  folderId?: string;
}

class SpoFolderRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Apply a retention label to a folder';
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
        folderId: typeof args.options.folderId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--name <name>'
      },
      {
        option: '--folderUrl [folderUrl]'
      },
      {
        option: 'i, --folderId [folderId]'
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
    try {
      const folderProperties = await this.getFolderProperties(logger, args);

      if (folderProperties.ListItemAllFields) {
        const options: SpoListItemRetentionLabelEnsureCommandOptions = {
          webUrl: args.options.webUrl,
          listId: folderProperties.ListItemAllFields.ParentList.Id,
          listItemId: folderProperties.ListItemAllFields.Id,
          name: args.options.name,
          output: 'json',
          debug: this.debug,
          verbose: this.verbose
        };

        const spoListItemRetentionLabelEnsureCommandOutput = await Cli.executeCommandWithOutput(spoListItemRetentionLabelEnsureCommand as Command, { options: { ...options, _: [] } });

        if (this.verbose) {
          await logger.logToStderr(spoListItemRetentionLabelEnsureCommandOutput.stderr);
        }
      }
      else {
        const options: SpoListRetentionLabelEnsureCommandOptions = {
          webUrl: args.options.webUrl,
          listUrl: folderProperties.ServerRelativeUrl,
          name: args.options.name,
          output: 'json',
          debug: this.debug,
          verbose: this.verbose
        };

        const spoListRetentionLabelEnsureCommandOutput = await Cli.executeCommandWithOutput(spoListRetentionLabelEnsureCommand as Command, { options: { ...options, _: [] } });

        if (this.verbose) {
          await logger.logToStderr(spoListRetentionLabelEnsureCommandOutput.stderr);
        }
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
      requestUrl += `GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<FolderProperties>(requestOptions);
  }
}

export default new SpoFolderRetentionLabelEnsureCommand();