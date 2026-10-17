import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileProperties } from './FileProperties.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { spo } from '../../../../utils/spo.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string(),
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional().alias('i'),
  assetId: z.string().optional().alias('a')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Applies a retention label to a file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileUrl, opts.fileId].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileUrl' or 'fileId', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileUrl', 'fileId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const fileProperties = await this.getFileProperties(logger, args);
      const parsedUrl = new URL(args.options.webUrl);
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