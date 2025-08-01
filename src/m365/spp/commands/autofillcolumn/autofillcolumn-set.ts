import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { ListInstance } from '../../../spo/commands/list/ListInstance.js';
import commands from '../../commands.js';
import { zod } from '../../../../utils/zod.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';

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

const options = globalOptionsZod
  .extend({
    siteUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))),
    listTitle: z.string().optional(),
    listId: z.string()
      .refine(value => validation.isValidGuid(value), listId => ({
        message: `${listId} in parameter listId is not a valid GUID`
      })).optional(),
    listUrl: z.string().optional(),
    columnId: zod.alias('i', z.string()
      .refine(value => validation.isValidGuid(value), columnId => ({
        message: `${columnId} in parameter columnId is not a valid GUID`
      })).optional()),
    columnTitle: zod.alias('t', z.string().optional()),
    prompt: z.string().optional(),
    isEnabled: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SppAutofillColumnSetCommand extends SpoCommand {
  public get name(): string {
    return commands.AUTOFILLCOLUMN_SET;
  }

  public get description(): string {
    return 'Applies the autofill option to the selected column';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.columnId, options.columnTitle].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'id' or 'title'.`
      })
      .refine(options => [options.listTitle, options.listId, options.listUrl].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'listTitle', 'listId' or 'listUrl'.`
      });
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

      if (!Object.values(AllowedFieldTypeKind).includes(column.FieldTypeKind)) {
        throw Error(`The specified column has incorrect type.`);
      }

      if (!!column.AutofillInfo) {
        await this.updateAutoFillColumnSettings(args, column.Id, listInstance.Id, column.AutofillInfo);
        return;
      }

      if (!args.options.prompt) {
        throw Error(`The prompt parameter is required when setting the autofill column for the first time.`);
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