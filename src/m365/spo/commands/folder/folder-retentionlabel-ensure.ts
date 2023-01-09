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
import { Options as SpoListItemRetentionLabelEnsureCommandOptions } from '../listitem/listitem-retentionlabel-ensure';
import * as SpoListItemRetentionLabelEnsureCommand from '../listitem/listitem-retentionlabel-ensure';
import { formatting } from '../../../../utils/formatting';

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
    if (this.verbose) {
      logger.logToStderr(`Applying retention label to folder ${args.options.folderId || args.options.folderUrl} in site at ${args.options.webUrl}...`);
    }
    try {
      const folderProperties = await this.getFolderProperties(args);
      const options: SpoListItemRetentionLabelEnsureCommandOptions = {
        webUrl: args.options.webUrl,
        listUrl: folderProperties.listServerRelativeUrl,
        listItemId: folderProperties.id,
        name: args.options.name,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      await Cli.executeCommandWithOutput(SpoListItemRetentionLabelEnsureCommand as Command, { options: { ...options, _: [] } });
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderProperties(args: CommandArgs): Promise<{ id: string, listServerRelativeUrl: string }> {
    const requestOptions: AxiosRequestConfig = {
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (args.options.folderId) {
      requestOptions.url = `${args.options.webUrl}/_api/web/GetFolderById('${args.options.folderId}')?$expand=ListItemAllFields`;
    }
    else {
      requestOptions.url = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(args.options.folderUrl!)}')?$expand=ListItemAllFields`;
    }

    const response = await request.get<FolderProperties>(requestOptions);
    return { id: response.ListItemAllFields.Id, listServerRelativeUrl: this.getListServerRelativeUrl(response.ServerRelativeUrl) };
  }

  private getListServerRelativeUrl(folderUrl: string): string {
    return folderUrl.replace(/\/[^\/]+$/, '');
  }
}

module.exports = new SpoFolderRetentionLabelEnsureCommand();