import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';

const options = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, {
        message: 'webUrl is not a valid SharePoint site URL.'
      })),
    listId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    listTitle: z.string().optional(),
    listUrl: z.string().optional(),
    title: z.string().min(1, 'Cannot be empty.'),
    fields: z.string().optional(),
    viewQuery: z.string().optional(),
    personal: z.boolean().optional(),
    default: z.boolean().optional(),
    paged: z.boolean().optional(),
    rowLimit: z.number().int().positive().optional(),
    customFormatter: z.string()
      .refine(formatter => {
        try {
          JSON.parse(formatter);
          return true;
        }
        catch {
          return false;
        }
      }, {
        message: 'Custom formatter must be a valid JSON string.'
      })
      .optional(),
    type: z.enum(['list', 'calendar', 'gallery', 'kanban']).optional(),
    calendarStartDateField: z.string().min(1, 'Cannot be empty.').optional(),
    calendarEndDateField: z.string().min(1, 'Cannot be empty.').optional(),
    calendarTitleField: z.string().min(1, 'Cannot be empty.').optional(),
    calendarSubTitleField: z.string().min(1, 'Cannot be empty.').optional(),
    calendarDefaultLayout: z.enum(['month', 'week', 'workWeek', 'day']).optional(),
    kanbanBucketField: z.string().min(1, 'Cannot be empty.').optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListViewAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_ADD;
  }

  public get description(): string {
    return 'Adds a new view to a SharePoint list';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine((options: Options) => [options.listId, options.listTitle, options.listUrl].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options: listId, listTitle, or listUrl.'
      })
      .refine((options: Options) => !options.personal || !options.default, {
        message: 'Default view cannot be a personal view.'
      })
      .refine((options: Options) => options.type !== 'calendar' || [options.calendarStartDateField, options.calendarEndDateField, options.calendarTitleField].filter(o => o === undefined).length === 0, {
        message: 'When type is calendar, do specify calendarStartDateField, calendarEndDateField, and calendarTitleField.'
      })
      .refine((options: Options) => options.type === 'calendar' || [options.calendarStartDateField, options.calendarEndDateField, options.calendarTitleField].filter(o => o === undefined).length === 3, {
        message: 'When type is not calendar, do not specify calendarStartDateField, calendarEndDateField, and calendarTitleField.'
      })
      .refine((options: Options) => options.type !== 'kanban' || options.kanbanBucketField !== undefined, {
        message: 'When type is kanban, do specify kanbanBucketField.'
      })
      .refine((options: Options) => options.type === 'kanban' || options.kanbanBucketField === undefined, {
        message: 'When type is not kanban, do not specify kanbanBucketField.'
      })
      .refine((options: Options) => options.type === 'calendar' || options.fields !== undefined, {
        message: 'When type is not calendar, do specify fields.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Adding view '${args.options.title}' to list...`);
      }

      let apiUrl = `${args.options.webUrl}/_api/web/`;
      if (args.options.listId) {
        apiUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        apiUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
      }
      else if (args.options.listUrl) {
        apiUrl += `GetList('${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl))}')`;
      }
      apiUrl += '/views/add';

      const requestOptions: CliRequestOptions = {
        url: apiUrl,
        headers: {
          'content-type': 'application/json;odata=verbose',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          parameters: {
            Title: args.options.title,
            ViewFields: {
              results: args.options.fields?.split(',').map(f => f.trim()) ?? []
            },
            Query: args.options.viewQuery,
            PersonalView: !!args.options.personal,
            SetAsDefaultView: !!args.options.default,
            Paged: !!args.options.paged,
            RowLimit: args.options.rowLimit ?? 30,
            CustomFormatter: args.options.customFormatter
          }
        }
      };

      this.setViewTypeSpecificParameters(args.options, requestOptions.data.parameters);

      const result = await request.post<any>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private setViewTypeSpecificParameters(options: Options, requestBody: any): void {
    if (options.type === 'calendar') {
      const defaultView = options.calendarDefaultLayout ?? 'month';
      const titleField = options.calendarTitleField!;
      const subTitleField = options.calendarSubTitleField ?? '';

      // Following fields are required for calendar view, order is important
      const viewFields = [options.calendarStartDateField, options.calendarEndDateField, titleField, subTitleField].filter(field => field !== '');

      // Add any additional fields specified by the user
      const extraViewFields = requestBody.ViewFields.results.filter((field: string) => !viewFields.includes(field.trim()));
      viewFields.push(...extraViewFields);

      requestBody.CalendarViewStyles = `<CalendarViewStyle Title="Day" Type="day" Template="CalendarViewdayChrome" Sequence="1" Default="${String(defaultView === 'day').toUpperCase()}" /><CalendarViewStyle Title="Week" Type="week" Template="CalendarViewweekChrome" Sequence="2" Default="${String(defaultView === 'week').toUpperCase()}" /><CalendarViewStyle Title="Month" Type="month" Template="CalendarViewmonthChrome" Sequence="3" Default="${String(defaultView === 'month').toUpperCase()}" /><CalendarViewStyle Title="Work week" Type="workweek" Template="CalendarViewweekChrome" Sequence="4" Default="${String(defaultView === 'workWeek').toUpperCase()}" />`;
      requestBody.Query = `<Where><DateRangesOverlap><FieldRef Name='${options.calendarStartDateField}' /><FieldRef Name='${options.calendarEndDateField}' /><Value Type='DateTime'><Month /></Value></DateRangesOverlap></Where>`;
      requestBody.ViewData = `<FieldRef Name="${titleField}" Type="CalendarMonthTitle" /><FieldRef Name="${titleField}" Type="CalendarWeekTitle" /><FieldRef Name="${subTitleField}" Type="CalendarWeekLocation" /><FieldRef Name="${titleField}" Type="CalendarDayTitle" /><FieldRef Name="${subTitleField}" Type="CalendarDayLocation" />`;
      requestBody.ViewFields.results = viewFields;
      requestBody.ViewType2 = 'MODERNCALENDAR';
      return;
    }
    if (options.type === 'gallery') {
      requestBody.ViewType2 = 'TILES';
      return;
    }
    if (options.type === 'kanban') {
      // Add the bucket field to the view fields if it is not already included
      const viewFields: string[] = requestBody.ViewFields.results;
      if (!viewFields.includes(options.kanbanBucketField!)) {
        viewFields.push(options.kanbanBucketField!);
      }

      if (!options.customFormatter) {
        requestBody.CustomFormatter = '{}';
      }

      requestBody.ViewData = `<FieldRef Name="${options.kanbanBucketField}" Type="KanbanPivotColumn" />`;
      requestBody.ViewType2 = 'KANBAN';
      return;
    }
  };
}

export default new SpoListViewAddCommand();