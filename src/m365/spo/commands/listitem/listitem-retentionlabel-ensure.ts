import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteRetentionLabel } from './SiteRetentionLabel';
import * as SpoWebRetentionLabelListCommand from '../web/web-retentionlabel-list';
import { Options as SpoWebRetentionLabelListCommandOptions } from '../web/web-retentionlabel-list';
import Command from '../../../../Command';
import { Cli } from '../../../../cli/Cli';
import { ListItemRetentionLabel } from './ListItemRetentionLabel';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  listItemId: string;
  name?: string;
  id?: string;
  assetId?: string;
}

class SpoListItemRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Apply a retention label to a list item';
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
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        id: typeof args.options.id !== 'undefined',
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
        option: '--listItemId <listItemId>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-a, --assetId [assetId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const id: number = parseInt(args.options.listItemId);
        if (isNaN(id)) {
          return `${args.options.listItemId} is not a valid list item ID`;
        }

        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId as string)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        if (args.options.id &&
          !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
    this.optionSets.push({ options: ['name', 'id'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const labelInformation = await this.getLabelInformation(args.options, logger);

      if (args.options.assetId) {
        await this.applyAssetId(args.options, logger);
      }

      await this.applyLabel(args.options, labelInformation, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getLabelInformation(options: Options, logger: Logger): Promise<ListItemRetentionLabel> {
    const cmdOptions: SpoWebRetentionLabelListCommandOptions = {
      webUrl: options.webUrl,
      output: 'json',
      debug: options.debug,
      verbose: options.verbose
    };

    const output = await Cli.executeCommandWithOutput(SpoWebRetentionLabelListCommand as Command, { options: { ...cmdOptions, _: [] } });

    if (this.verbose) {
      logger.logToStderr(output.stderr);
    }

    const labels = JSON.parse(output.stdout) as SiteRetentionLabel[];
    const label = labels.find(l => l.TagName === options.name || l.TagId === options.id);

    if (this.verbose && label !== undefined) {
      logger.logToStderr(`Retention label found in the list of available labels: '${label.TagName}' / '${label.TagId}'...`);
    }

    if (label === undefined) {
      throw new Error(`The specified retention label does not exist`);
    }

    return {
      complianceTag: label.TagName,
      isTagPolicyHold: label.BlockDelete,
      isTagPolicyRecord: label.BlockEdit,
      isEventBasedTag: label.IsEventTag,
      isTagSuperLock: label.SuperLock,
      isUnlockedAsDefault: label.UnlockedAsDefault
    } as ListItemRetentionLabel;
  }

  private async applyLabel(options: Options, labelInformation: ListItemRetentionLabel, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Applying retention label to item in list '${options.listId || options.listTitle || options.listUrl}' in site at ${options.webUrl}...`);
    }

    let requestUrl = `${options.webUrl}/_api/web`;

    if (options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')/items(${options.listItemId})/SetComplianceTag()`;
    }
    else if (options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/items(${options.listItemId})/SetComplianceTag()`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      requestUrl += `/GetList(@listUrl)/items(${options.listItemId})/SetComplianceTag()?@listUrl='${formatting.encodeQueryParameter(listServerRelativeUrl)}'`;
    }

    const requestOptions: AxiosRequestConfig = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: labelInformation,
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  async applyAssetId(options: GlobalOptions, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Applying the asset Id ${options.assetId}...`);
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

module.exports = new SpoListItemRetentionLabelEnsureCommand();