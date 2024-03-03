import * as url from 'url';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileProperties } from './FileProperties.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { spo } from '../../../../utils/spo.js';

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
    this.#initTypes();
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

  #initTypes(): void {
    this.types.string.push('webUrl', 'name', 'fileUrl', 'fileId', 'assetId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const fileProperties = await this.getFileProperties(logger, args);
      const parsedUrl = url.parse(args.options.webUrl);
      const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
      const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, fileProperties.listServerRelativeUrl);

      if (args.options.assetId) {
        await this.applyAssetId(args.options.webUrl, fileProperties.listServerRelativeUrl, fileProperties.listItemId, args.options.assetId);
      }

      await spo.applyRetentionLabelToListItems(args.options.webUrl, args.options.name, listAbsoluteUrl, [parseInt(fileProperties.listItemId)], logger, args.options.verbose);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFileProperties(logger: Logger, args: CommandArgs): Promise<{ listItemId: string, listServerRelativeUrl: string }> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list and item information for file '${args.options.fileId || args.options.fileUrl}' in site at ${args.options.webUrl}...`);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.fileId) {
      requestUrl += `GetFileById('${args.options.fileId}')`;
    }
    else {
      const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl!);
      requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<FileProperties>(requestOptions);
    return { listItemId: response.ListItemAllFields.Id, listServerRelativeUrl: response.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl };
  }

  private async applyAssetId(webUrl: string, listServerRelativeUrl: string, listItemId: string, assetId: string): Promise<void> {
    const requestBody = { "formValues": [{ "FieldName": "ComplianceAssetId", "FieldValue": assetId }] };

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetList(@listUrl)/items(${listItemId})/ValidateUpdateListItem()?@listUrl='${formatting.encodeQueryParameter(listServerRelativeUrl)}'`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    return request.post(requestOptions);
  }
}

export default new SpoFileRetentionLabelEnsureCommand();