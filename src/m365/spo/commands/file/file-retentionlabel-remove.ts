import { AxiosRequestConfig } from 'axios';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListItemRetentionLabelRemoveCommand, { Options as SpoListItemRetentionLabelRemoveCommandOptions } from '../listitem/listitem-retentionlabel-remove.js';
import { FileProperties } from './FileProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  force?: boolean;
}

class SpoFileRetentionLabelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_RETENTIONLABEL_REMOVE;
  }

  public get description(): string {
    return 'Clear the retention label from a file';
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
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        fileId: typeof args.options.fileId !== 'undefined',
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
        option: '--fileUrl [fileUrl]'
      },
      {
        option: '-i, --fileId [fileId]'
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

        if (args.options.fileId &&
          !validation.isValidGuid(args.options.fileId as string)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeFileRetentionLabel(logger, args);
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the retentionlabel from file ${args.options.fileId || args.options.fileUrl} located in site ${args.options.webUrl}?` });

      if (result) {
        await this.removeFileRetentionLabel(logger, args);
      }
    }
  }

  private async removeFileRetentionLabel(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing retention label from file ${args.options.fileId || args.options.fileUrl} in site at ${args.options.webUrl}...`);
    }
    try {
      const fileProperties = await this.getFileProperties(args);
      const options: SpoListItemRetentionLabelRemoveCommandOptions = {
        webUrl: args.options.webUrl,
        listId: fileProperties.listId,
        listItemId: fileProperties.id,
        force: true,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      const spoListItemRetentionLabelRemoveCommandOutput = await Cli.executeCommandWithOutput(spoListItemRetentionLabelRemoveCommand as Command, { options: { ...options, _: [] } });
      if (this.verbose) {
        await logger.logToStderr(spoListItemRetentionLabelRemoveCommandOutput.stderr);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFileProperties(args: CommandArgs): Promise<{ id: string, listId: string }> {
    const requestOptions: AxiosRequestConfig = {
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.fileId) {
      requestUrl += `GetFileById('${args.options.fileId}')`;
    }
    else {
      const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl!);
      requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }

    requestOptions.url = `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ListItemAllFields/ParentList/Id,ListItemAllFields/Id`;

    const response = await request.get<FileProperties>(requestOptions);
    return { id: response.ListItemAllFields.Id, listId: response.ListItemAllFields.ParentList.Id };
  }
}

export default new SpoFileRetentionLabelRemoveCommand();