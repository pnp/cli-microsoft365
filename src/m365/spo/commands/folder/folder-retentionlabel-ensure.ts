import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderProperties } from './FolderProperties';
import { Options as SpoListItemRetentionLabelEnsureCommandOptions } from '../listitem/listitem-retentionlabel-ensure';
import * as SpoListItemRetentionLabelEnsureCommand from '../listitem/listitem-retentionlabel-ensure';
import { Options as SpoListRetentionLabelEnsureCommandOptions } from '../list/list-retentionlabel-ensure';
import * as SpoListRetentionLabelEnsureCommand from '../list/list-retentionlabel-ensure';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';

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

        const spoListItemRetentionLabelEnsureCommandOutput = await Cli.executeCommandWithOutput(SpoListItemRetentionLabelEnsureCommand as Command, { options: { ...options, _: [] } });

        if (this.verbose) {
          logger.logToStderr(spoListItemRetentionLabelEnsureCommandOutput.stderr);
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

        const spoListRetentionLabelEnsureCommandOutput = await Cli.executeCommandWithOutput(SpoListRetentionLabelEnsureCommand as Command, { options: { ...options, _: [] } });

        if (this.verbose) {
          logger.logToStderr(spoListRetentionLabelEnsureCommandOutput.stderr);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderProperties(logger: Logger, args: CommandArgs): Promise<FolderProperties> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list and item information for folder '${args.options.folderId || args.options.folderUrl}' in site at ${args.options.webUrl}...`);
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

module.exports = new SpoFolderRetentionLabelEnsureCommand();