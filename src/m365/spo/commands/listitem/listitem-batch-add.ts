import { AxiosRequestConfig } from 'axios';
import * as fs from 'fs';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { v4 } from 'uuid';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
}

interface FormValues {
  FieldName: string;
  FieldValue: string;
}

class SpoListItemBatchAddCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_BATCH_ADD;
  }

  public get description(): string {
    return 'Creates list items in a batch';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --filePath <filePath>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
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

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        if (!fs.existsSync(args.options.filePath)) {
          return `File with path ${args.options.filePath} does not exist`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push(
      'webUrl',
      'filePath',
      'listId',
      'listTitle',
      'listUrl'
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Starting to create batch items from csv at path ${args.options.filePath}`);
      }
      const csvContent = fs.readFileSync(args.options.filePath, 'utf8');
      const jsonContent = formatting.parseCsvToJson(csvContent);
      await this.addItemsAsBatch(jsonContent, args.options, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addItemsAsBatch(rows: any[], options: Options, logger: Logger): Promise<void> {
    const requestUrl = this.getRequestUrl(options);
    let itemsToAdd: FormValues[][] = [];

    for await (const [index, row] of rows.entries()) {
      itemsToAdd.push(this.getSingleItemRequestBody(row));
      if (itemsToAdd.length === 100) {
        if (this.verbose) {
          logger.logToStderr(`Writing away batch of items, currently at: ${index + 1}/${rows.length}.`);
        }
        await this.postBatchData(itemsToAdd, options.webUrl, requestUrl);
        itemsToAdd = [];
      }
    }
    if (itemsToAdd.length) {
      if (this.verbose) {
        logger.logToStderr(`Writing away ${itemsToAdd.length} items.`);
      }
      await this.postBatchData(itemsToAdd, options.webUrl, requestUrl);
    }
  }

  private async postBatchData(itemsToAdd: FormValues[][], webUrl: string, requestUrl: string): Promise<void> {
    const batchId = v4();
    const requestBody = this.parseBatchRequestBody(itemsToAdd, batchId, requestUrl);
    const requestOptions: AxiosRequestConfig = {
      url: `${webUrl}/_api/$batch`,
      headers: {
        'Content-Type': `multipart/mixed; boundary=batch_${batchId}`,
        'Accept': 'application/json;odata=verbose'
      },
      data: requestBody.join('')
    };
    await request.post(requestOptions);
  }

  private parseBatchRequestBody(items: FormValues[][], batchId: string, requestUrl: string): string[] {
    const changeSetId = v4();
    const batchBody: string[] = [];

    // add default batch body headers
    batchBody.push(`--batch_${batchId}\n`);
    batchBody.push(`Content-Type: multipart/mixed; boundary="changeset_${changeSetId}"\n\n`);
    batchBody.push('Content-Transfer-Encoding: binary\n\n');

    items.forEach((item) => {
      batchBody.push(`--changeset_${changeSetId}\n`);
      batchBody.push('Content-Type: application/http\n');
      batchBody.push('Content-Transfer-Encoding: binary\n\n');
      batchBody.push(`POST ${requestUrl} HTTP/1.1\n`);
      batchBody.push(`Accept: application/json;odata=nometadata\n`);
      batchBody.push(`Content-Type: application/json;odata=verbose\n`);
      batchBody.push(`If-Match: *\n\n`);
      batchBody.push(`{\n"formValues": ${JSON.stringify(item)}\n}`);
    });

    // close batch body
    batchBody.push(`\n\n`);
    batchBody.push(`--changeset_${changeSetId}--\n\n`);
    batchBody.push(`--batch_${batchId}--\n`);

    return batchBody;
  }

  private getSingleItemRequestBody(row: any): FormValues[] {
    const requestBody: FormValues[] = [];
    Object.keys(row).forEach(key => {
      // have to do 'toString()' or the API will complain when entering a numeric field
      requestBody.push({ FieldName: key, FieldValue: (<any>row)[key].toString() });
    });
    return requestBody;
  }

  private getRequestUrl(options: Options): string {
    let listUrl = `${options.webUrl}/_api/web`;
    if (options.listId) {
      listUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
    }
    else if (options.listTitle) {
      listUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      listUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    return `${listUrl}/AddValidateUpdateItemUsingPath`;
  }
}

module.exports = new SpoListItemBatchAddCommand();