import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { ListInstance } from '../../../spo/commands/list/ListInstance.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import z from 'zod';

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
  GridChoice = 16
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

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string().refine(url =>
    validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listTitle: z.string().optional(),
  listId: z.string().uuid()
    .refine(value => validation.isValidGuid(value), {
      error: e => `'${e.input}' in parameter listId is not a valid GUID.`
    }).optional(),
  listUrl: z.string().optional(),
  columnId: z.string().uuid()
    .refine(value => validation.isValidGuid(value), {
      error: e => `'${e.input}' in parameter columnId is not a valid GUID.`
    }).optional().alias('i'),
  columnTitle: z.string().optional().alias('t'),
  columnInternalName: z.string().optional(),
  prompt: z.string().optional(),
  isEnabled: z.boolean().optional()
}).strict();

declare type Options = z.infer<typeof options>;

class SppAutofillColumnSetCommand extends SpoCommand {
  public get name(): string {
    return commands.AUTOFILLCOLUMN_SET;
  }

  public get description(): string {
    return 'Applies the autofill option to the selected column';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.columnId, options.columnTitle, options.columnInternalName].filter(Boolean).length === 1, {
        message: `Specify exactly one of the following options: 'columnId', 'columnTitle' or 'columnInternalName'.`
      })
      .refine(options => [options.listTitle, options.listId, options.listUrl].filter(Boolean).length === 1, {
        message: `Specify exactly one of the following options: 'listTitle', 'listId' or 'listUrl'.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.log(`Applying an autofill column to a column...`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      const listInstance = await this.getDocumentLibraryInfo(siteUrl, args.options);

      if (listInstance.BaseType !== 1) {
        throw Error(`The specified list is not a document library.`);
      }

      const column = await this.getColumn(siteUrl, args.options, listInstance.Id);

      if (!Object.values(AllowedFieldTypeKind).includes(column.FieldTypeKind)) {
        throw Error(`The specified column has incorrect type.`);
      }

      if (column.AutofillInfo) {
        await this.updateAutoFillColumnSettings(siteUrl, args.options, column.Id, listInstance.Id, column.AutofillInfo);
        return;
      }

      if (!args.options.prompt) {
        throw Error(`The prompt parameter is required when setting the autofill column for the first time.`);
      }

      await this.applyAutoFillColumnSettings(siteUrl, args.options, column.Id, column.Title, listInstance.Id);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getDocumentLibraryInfo(siteUrl: string, options: Options): Promise<ListInstance> {
    let requestUrl = `${siteUrl}/_api/web`;

    if (options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
    }
    else if (options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(siteUrl, options.listUrl);
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

  private getColumn(siteUrl: string, options: Options, listId: string): Promise<Field> {
    let fieldRestUrl: string = '';

    if (options.columnId) {
      fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(options.columnId)}')`;
    }
    else {
      fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter((options.columnTitle || options.columnInternalName) as string)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields${fieldRestUrl}?&$select=Id,Title,FieldTypeKind,AutofillInfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get<Field>(requestOptions);
  }

  private updateAutoFillColumnSettings(siteUrl: string, options: Options, columnId: string, listInstanceId: string, autofillInfo: string): Promise<any> {
    const autofillInfoObj = JSON.parse(autofillInfo) as AutofillInfo;

    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/machinelearning/SetColumnLLMInfo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        autofillPrompt: options.prompt ?? autofillInfoObj.LLM.Prompt,
        columnId: columnId,
        docLibId: `{${listInstanceId}}`,
        isEnabled: options.isEnabled !== undefined ? options.isEnabled : autofillInfoObj.LLM.IsEnabled
      }
    };

    return request.post(requestOptions);
  }

  private applyAutoFillColumnSettings(siteUrl: string, options: Options, columnId: string, columnTitle: string, listInstanceId: string): Promise<any> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/machinelearning/SetSyntexPoweredColumnPrompts`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        docLibId: `{${listInstanceId}}`,
        syntexPoweredColumnPrompts: JSON.stringify([{
          columnId: columnId,
          columnName: columnTitle,
          prompt: options.prompt,
          isEnabled: options.isEnabled !== undefined ? options.isEnabled : true
        }])
      }
    };

    return request.post(requestOptions);
  }
}

export default new SppAutofillColumnSetCommand();