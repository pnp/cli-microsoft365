import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileProperties } from './FileProperties';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { spo } from '../../../../utils/spo';
import { Options as SpoListItemRetentionLabelEnsureCommandOptions } from '../listitem/listitem-retentionlabel-ensure';
import * as SpoListItemRetentionLabelEnsureCommand from '../listitem/listitem-retentionlabel-ensure';
import { ListItemRetentionLabel } from '../listitem/ListItemRetentionLabel';
import { Cli } from '../../../../cli/Cli';
import Command from '../../../../Command';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  fileUrl?: string;
  fileId?: string;
  assetId?: string;
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
        fileId: typeof args.options.fileId !== 'undefined',
        assetId: typeof args.options.assetId !== 'undefined'
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
      },
      {
        option: '-a, --assetId [assetId]'
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
      const fileProperties: FileProperties = await this.getFileProperties(logger, args);

      const labelInformation: ListItemRetentionLabel = await spo.getWebRetentionLabelInformationByName(args.options.webUrl, args.options.name);

      if (args.options.assetId && !labelInformation.isEventBasedTag) {
        throw `The label that's being applied is not an event-based label`;
      }
      if (args.options.assetId && labelInformation.isEventBasedTag) {
        await this.applyAssetId(args.options.webUrl, fileProperties.ListItemAllFields.ParentList.Id, fileProperties.ListItemAllFields.Id, args.options.assetId);
      }

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

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<FileProperties>(requestOptions);
  }

  private async applyAssetId(webUrl: string, listId: string, listItemId: string, assetId: string): Promise<void> {
    const requestUrl = `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')`;

    const requestBody = { "formValues": [{ "FieldName": "ComplianceAssetId", "FieldValue": assetId }] };

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}/items(${listItemId})/ValidateUpdateListItem()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    await request.post(requestOptions);
  }
}

module.exports = new SpoFileRetentionLabelEnsureCommand();