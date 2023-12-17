import fs from 'fs';
import os from 'os';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemFieldValueResult } from './ListItemFieldValueResult.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  filePath?: string;
  csvContent?: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
}

interface FormValues {
  FieldName: string;
  FieldValue: string;
}

interface BatchResult extends ListItemFieldValueResult {
  csvLineNumber: number
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
        filePath: typeof args.options.filePath !== 'undefined',
        csvContent: typeof args.options.csvContent !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-p, --filePath [filePath]'
      },
      {
        option: '-c, --csvContent [csvContent]'
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

        if (args.options.filePath && !fs.existsSync(args.options.filePath)) {
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
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['filePath', 'csvContent'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Starting to create batch items from csv ${args.options.filePath ? `at path ${args.options.filePath}` : 'from content'}`);
      }
      const csvContent = args.options.filePath ? fs.readFileSync(args.options.filePath, 'utf8') : args.options.csvContent;
      const jsonContent = formatting.parseCsvToJson(csvContent!);
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
          await logger.logToStderr(`Writing away batch of items, currently at: ${index + 1}/${rows.length}.`);
        }

        await this.postBatchData(itemsToAdd, options.webUrl, requestUrl);

        itemsToAdd = [];
      }
    }
    if (itemsToAdd.length) {
      if (this.verbose) {
        await logger.logToStderr(`Writing away ${itemsToAdd.length} items.`);
      }

      await this.postBatchData(itemsToAdd, options.webUrl, requestUrl);
    }
  }

  private async postBatchData(itemsToAdd: FormValues[][], webUrl: string, requestUrl: string): Promise<void> {
    const batchId = v4();
    const requestBody = this.parseBatchRequestBody(itemsToAdd, batchId, requestUrl);
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/$batch`,
      headers: {
        'Content-Type': `multipart/mixed; boundary=batch_${batchId}`,
        'Accept': 'application/json;odata=verbose'
      },
      data: requestBody.join('')
    };
    const response: any = await request.post(requestOptions);
    const parsedResponse = this.parseBatchResponseBody(response);

    if (parsedResponse.some(r => r.HasException)) {
      throw `Creating some items failed with the following errors: ${os.EOL}${parsedResponse.filter(f => f.HasException).map(f => { return `- Line ${f.csvLineNumber}: ${f.FieldName} - ${f.ErrorMessage}`; }).join(os.EOL)}`;
    }
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
      batchBody.push(`{\n"formValues": ${JSON.stringify(this.formatFormValues(item))}\n}`);
    });

    // close batch body
    batchBody.push(`\n\n`);
    batchBody.push(`--changeset_${changeSetId}--\n\n`);
    batchBody.push(`--batch_${batchId}--\n`);

    return batchBody;
  }

  private formatFormValues(input: FormValues[]): FormValues[] {
    // Fix for PS 7
    const output: FormValues[] = input.map(obj => ({
      FieldName: obj.FieldName.replace(/\\"/g, ''),
      FieldValue: obj.FieldValue.replace(/\\"/g, '')
    }));

    return output;
  }

  private parseBatchResponseBody(response: string): BatchResult[] {
    const batchResults: BatchResult[] = [];

    response.split('\r\n')
      .filter((line: string) => line.startsWith('{'))
      .forEach((line: string, index: number) => {
        const parsedResponse: any = JSON.parse(line);

        if (parsedResponse.error) {
          // if an error object is returned, the request failed
          const error = parsedResponse.error as { message: { value: string } };
          throw error.message.value;
        }

        (parsedResponse as { value: ListItemFieldValueResult[] }).value.forEach((fieldValueResult: ListItemFieldValueResult) => {
          batchResults.push({
            csvLineNumber: (index + 2),
            ...fieldValueResult
          } as BatchResult);
        });
      });

    return batchResults;
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

export default new SpoListItemBatchAddCommand();