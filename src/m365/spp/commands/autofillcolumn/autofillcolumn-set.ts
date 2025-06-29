import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { ListInstance } from '../../../spo/commands/list/ListInstance.js';
import commands from '../../commands.js';

enum AllowedFieldTypeKind {
  Integer = 1,
  Text = 2,
  Note = 3,
  DateTime = 4,
  Counter = 5,
  Choice = 6,
  Boolean = 8,
  Number = 9,
  Currency = 10,
  URL = 11,
  Computed = 12,
  MultiChoice = 15,
  GridChoice = 16,
}

interface Field {
  Id: string;
  Title: string;
  FieldTypeKind: number;
  AutofillInfo?: string;
}

interface AutofillInfo {
  LLM: {
    Prompt: string;
    IsEnabled: boolean;
  }
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  listTitle?: string;
  listId?: string;
  listUrl?: string;
  columnId?: string;
  columnTitle?: string;
  prompt?: string;
  isEnabled?: boolean;
}

class SppAutofillColumnSetCommand extends SpoCommand {
  public get name(): string {
    return commands.AUTOFILLCOLUMN_SET;
  }

  public get description(): string {
    return 'Applies the autofill option to the selected column';
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
        listTitle: typeof args.options.listTitle !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        columnId: typeof args.options.columnId !== 'undefined',
        columnTitle: typeof args.options.columnTitle !== 'undefined',
        prompt: typeof args.options.prompt !== 'undefined',
        isEnabled: !!args.options.isEnabled
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-i, --columnId [columnId]'
      },
      {
        option: '-t, --columnTitle [columnTitle]'
      },
      {
        option: '--prompt [prompt]'
      },
      {
        option: '--isEnabled [isEnabled]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.columnId && !validation.isValidGuid(args.options.columnId)) {
          return `${args.options.columnId} in parameter columnId is not a valid GUID`;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in parameter listId is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listTitle', 'listId', 'listUrl'] });
    this.optionSets.push({ options: ['columnTitle', 'columnId'] });
  }

  #initTypes(): void {
    this.types.string.push('siteUrl', 'listTitle', 'listId', 'listUrl', 'columnId', 'columnTitle', 'prompt');
    this.types.boolean.push('isEnabled');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.log(`Applying an autofill column to a column...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      const listInstance = await this.getDocumentLibraryInfo(args);

      if (listInstance.BaseType !== 1) {
        throw Error(`The specified list is not a document library.`);
      }

      const column = await this.getColumn(args.options, listInstance.Id);

      if (!(column.FieldTypeKind in AllowedFieldTypeKind)) {
        throw Error(`The specified column has incorrect type.`);
      }

      if (!!column.AutofillInfo) {
        await this.updateAutoFillColumnSettings(args, column.Id, listInstance.Id, column.AutofillInfo);
        return;
      }

      if (!args.options.prompt) {
        throw Error(`The prompt parameter is required for the first time setting the autofill column.`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${siteUrl}/_api/machinelearning/SetSyntexPoweredColumnPrompts`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: {
          docLibId: `{${listInstance.Id}}`,
          syntexPoweredColumnPrompts: JSON.stringify([{
            columnId: column.Id,
            columnName: column.Title,
            prompt: args.options.prompt,
            isEnabled: args.options.isEnabled !== undefined ? args.options.isEnabled : true
          }])
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getDocumentLibraryInfo(args: CommandArgs): Promise<ListInstance> {
    let requestUrl = `${args.options.siteUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.siteUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$select=Id,BaseType`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get<ListInstance>(requestOptions);
  }

  private getColumn(options: Options, listId: string): Promise<Field> {
    let fieldRestUrl: string = '';
    if (options.columnId) {
      fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(options.columnId)}')`;
    }
    else {
      fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter(options.columnTitle!)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${options.siteUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields${fieldRestUrl}?&$select=Id,Title,FieldTypeKind,AutofillInfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    return request.get<Field>(requestOptions);
  }

  private updateAutoFillColumnSettings(args: CommandArgs, columnId: string, listInstanceId: string, autofillInfo: string): Promise<any> {
    const autofillInfoObj = JSON.parse(autofillInfo) as AutofillInfo;

    const requestOptions: CliRequestOptions = {
      url: `${args.options.siteUrl}/_api/machinelearning/SetColumnLLMInfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        autofillPrompt: !!args.options.prompt ? args.options.prompt : autofillInfoObj.LLM.Prompt,
        columnId: columnId,
        docLibId: `{${listInstanceId}}`,
        isEnabled: args.options.isEnabled !== undefined ? args.options.isEnabled : autofillInfoObj.LLM.IsEnabled
      }
    };

    return request.post(requestOptions);
  }
}

export default new SppAutofillColumnSetCommand();