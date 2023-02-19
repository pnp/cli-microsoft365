import { AxiosRequestConfig } from 'axios';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileProperties } from './FileProperties';
import { Options as SpoListItemRetentionLabelEnsureCommandOptions } from '../listitem/listitem-retentionlabel-ensure';
import * as SpoListItemRetentionLabelEnsureCommand from '../listitem/listitem-retentionlabel-ensure';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  fileUrl?: string;
  fileId?: string;
}

class SpoFileRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Apply a retention label to a file';
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
        fileId: typeof args.options.fileId !== 'undefined'
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
        option: '--fileUrl [fileUrl]'
      },
      {
        option: '-i, --fileId [fileId]'
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
    try {
      const fileProperties = await this.getFileProperties(logger, args);

      const options: SpoListItemRetentionLabelEnsureCommandOptions = {
        webUrl: args.options.webUrl,
        listId: fileProperties.ListItemAllFields.ParentList.Id,
        listItemId: fileProperties.ListItemAllFields.Id,
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
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFileProperties(logger: Logger, args: CommandArgs): Promise<FileProperties> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list and item information for file '${args.options.fileId || args.options.fileUrl}' in site at ${args.options.webUrl}...`);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.fileId) {
      requestUrl += `GetFileById('${args.options.fileId}')`;
    }
    else {
      const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl!);
      requestUrl += `GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<FileProperties>(requestOptions);
  }
}

module.exports = new SpoFileRetentionLabelEnsureCommand();