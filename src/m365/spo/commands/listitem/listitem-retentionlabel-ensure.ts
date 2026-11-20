import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { SiteRetentionLabel } from './SiteRetentionLabel.js';
import { odata } from '../../../../utils/odata.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listItemId: z.number(),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  name: z.string().optional().alias('n'),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional()
    .alias('i'),
  assetId: z.string().optional().alias('a')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Applies a retention label to a list item';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify either listId, listTitle or listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      })
      .refine(opts => [opts.name, opts.id].filter(o => o !== undefined).length === 1, {
        error: 'Specify either name or id.',
        params: {
          customCode: 'optionSet',
          options: ['name', 'id']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const listAbsoluteUrl = await this.getListAbsoluteUrl(args.options, logger);
      const labelName = await this.getLabelName(args.options, logger);

      if (args.options.assetId) {
        await this.applyAssetId(args.options, logger);
      }

      await spo.applyRetentionLabelToListItems(args.options.webUrl, labelName, listAbsoluteUrl, [args.options.listItemId], logger, args.options.verbose);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }


  private async getLabelName(options: Options, logger: Logger): Promise<string> {
    if (options.name) {
      return options.name;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving the name of the retention label based on the Id '${options.id}'...`);
    }

    const requestUrl: string = `${options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(options.webUrl)}'`;
    const labels = await odata.getAllItems<SiteRetentionLabel>(requestUrl);
    const label = labels.find(l => l.TagId === options.id);

    if (label === undefined) {
      throw new Error(`The specified retention label does not exist or is not published to this SharePoint site. Use the name of the label if you want to apply an unpublished label.`);
    }

    if (this.verbose && label !== undefined) {
      await logger.logToStderr(`Retention label found in the list of available labels: '${label.TagName}' / '${label.TagId}'...`);
    }

    return label.TagName;
  }

  private async getListAbsoluteUrl(options: Options, logger: Logger): Promise<string> {
    const parsedUrl = new URL(options.webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    if (options.listUrl) {
      const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      return urlUtil.urlCombine(tenantUrl, serverRelativePath);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving list absolute URL...`);
    }

    let requestUrl = `${options.webUrl}/_api/web`;

    if (options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
    }
    else if (options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ RootFolder: { ServerRelativeUrl: string } }>(requestOptions);
    const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, response.RootFolder.ServerRelativeUrl);
    const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, serverRelativePath);

    if (this.verbose) {
      await logger.logToStderr(`List absolute URL found: '${listAbsoluteUrl}'`);
    }

    return listAbsoluteUrl;
  }

  async applyAssetId(options: Options, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Applying the asset Id ${options.assetId}...`);
    }

    let requestUrl = `${options.webUrl}/_api/web`;

    if (options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')/items(${options.listItemId})/ValidateUpdateListItem()`;
    }
    else if (options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/items(${options.listItemId})/ValidateUpdateListItem()`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      requestUrl += `/GetList(@listUrl)/items(${options.listItemId})/ValidateUpdateListItem()?@listUrl='${formatting.encodeQueryParameter(listServerRelativeUrl)}'`;
    }

    const requestBody = { "formValues": [{ "FieldName": "ComplianceAssetId", "FieldValue": options.assetId }] };

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    await request.post(requestOptions);
  }
}

export default new SpoListItemRetentionLabelEnsureCommand();