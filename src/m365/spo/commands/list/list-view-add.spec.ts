import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './list-view-add.js';

describe(commands.LIST_VIEW_ADD, () => {

  const validListTitle = 'List title';
  const validListId = '00000000-0000-0000-0000-000000000000';
  const validListUrl = '/Lists/SampleList';
  const validTitle = 'View title';
  const validWebUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const validFieldsInput = 'Field1,Field2,Field3';

  const viewCreationResponse = {
    DefaultView: false,
    Hidden: false,
    Id: "00000000-0000-0000-0000-000000000000",
    MobileDefaultView: false,
    MobileView: false,
    Paged: true,
    PersonalView: false,
    ViewProjectedFields: null,
    ViewQuery: "",
    RowLimit: 30,
    Scope: 0,
    ServerRelativePath: {
      DecodedUrl: `/sites/project-x/Lists/${validListTitle}/${validTitle}.aspx`
    },
    ServerRelativeUrl: `/sites/project-x/Lists/${validListTitle}/${validTitle}.aspx`,
    Title: validTitle
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_VIEW_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'invalid',
      listTitle: validListTitle,
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if listId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listId: 'invalid',
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if rowLimit is not a number', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput,
      rowLimit: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if rowLimit is lower than 1', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput,
      rowLimit: 0
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when setting default and personal option', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput,
      personal: true,
      default: true
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when formatting is not a valid JSON string', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput,
      customFormatter: 'invalid json'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails when listId and listTitle are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listId: validListId,
      listTitle: validListTitle,
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails when listUrl and listTitle are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listTitle: validListTitle,
      listUrl: validListUrl,
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails when not listId, listTitle, nor listUrl are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when calendarStartDateField is not specified for a calendar type', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      type: 'calendar',
      calendarEndDateField: 'EndDate',
      calendarTitleField: 'Title',
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when calendarEndDateField is not specified for a calendar type', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      type: 'calendar',
      calendarStartDateField: 'StartDate',
      calendarTitleField: 'Title',
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when calendarTitleField is not specified for a calendar type', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      type: 'calendar',
      calendarStartDateField: 'StartDate',
      calendarEndDateField: 'EndDate',
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when calendar fields are specified for a regular view', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      calendarStartDateField: 'StartDate',
      calendarEndDateField: 'EndDate',
      calendarTitleField: 'Title',
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when creating a kanban view without the kanban field', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      type: 'kanban',
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when kanban fields are specified for a regular view', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      kanbanBucketField: 'Bucket',
      title: validTitle,
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when fields field is not specified and type is not calendar', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      title: validTitle
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when specifying an incorrect type', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      title: validTitle,
      type: 'invalid',
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when specifying an incorrect calendarDefaultLayout', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      title: validTitle,
      calendarDefaultLayout: 'invalid',
      fields: validFieldsInput
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('correctly validates regular view options', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly validates gallery view options', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      type: 'gallery',
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly validates calendar view options', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      type: 'calendar',
      calendarStartDateField: 'StartDate',
      calendarEndDateField: 'EndDate',
      calendarTitleField: 'Title',
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput,
      customFormatter: JSON.stringify({ someProperty: 'someValue' })
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly validates kanban view options', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      type: 'kanban',
      kanbanBucketField: 'Bucket',
      listId: validListId,
      title: validTitle,
      fields: validFieldsInput,
      customFormatter: JSON.stringify({ someProperty: 'someValue' })
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly sets default paged value when paged option is not specified', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(validListTitle)}')/views/add`) {
        return viewCreationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listTitle: validListTitle,
        title: validTitle,
        fields: validFieldsInput
      }
    });

    // Verify that Paged defaults to true when not specified
    assert.strictEqual(postStub.lastCall.args[0].data.parameters.Paged, true);
  });

  it('correctly logs an output', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(validListTitle)}')/views/add`) {
        return viewCreationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listTitle: validListTitle,
        title: validTitle,
        fields: validFieldsInput
      }
    });
    assert(loggerLogSpy.calledWith(viewCreationResponse));
  });

  it('correctly creates a regular view by list title', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(validListTitle)}')/views/add`) {
        return viewCreationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listTitle: validListTitle,
        title: validTitle,
        fields: validFieldsInput
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data.parameters, {
      Title: validTitle,
      ViewFields: {
        results: validFieldsInput.split(',')
      },
      CustomFormatter: undefined,
      Query: undefined,
      PersonalView: false,
      SetAsDefaultView: false,
      Paged: true,
      RowLimit: 30
    });
  });

  it('correctly adds a kanban view by list id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(validListId)}')/views/add`) {
        return viewCreationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listId: validListId,
        type: 'kanban',
        title: validTitle,
        fields: validFieldsInput,
        kanbanBucketField: 'Status'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data.parameters, {
      Title: validTitle,
      ViewFields: {
        results: [...validFieldsInput.split(','), 'Status']
      },
      CustomFormatter: '{}',
      Query: undefined,
      ViewData: '<FieldRef Name="Status" Type="KanbanPivotColumn" />',
      PersonalView: false,
      SetAsDefaultView: false,
      Paged: true,
      RowLimit: 30,
      ViewType2: 'KANBAN'
    });
  });

  it('correctly adds gallery view by list URL', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetList('${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(validWebUrl, validListUrl))}')/views/add`) {
        return viewCreationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        type: 'gallery',
        listUrl: validListUrl,
        title: validTitle,
        fields: validFieldsInput,
        rowLimit: 100,
        verbose: true
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data.parameters, {
      Title: validTitle,
      ViewFields: {
        results: validFieldsInput.split(',')
      },
      CustomFormatter: undefined,
      Query: undefined,
      PersonalView: false,
      SetAsDefaultView: false,
      Paged: true,
      RowLimit: 100,
      ViewType2: 'TILES'
    });
  });

  it('correctly adds a calendar view by list id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(validListId)}')/views/add`) {
        return viewCreationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listId: validListId,
        type: 'calendar',
        title: validTitle,
        calendarStartDateField: 'StartDate',
        calendarEndDateField: 'EndDate',
        calendarTitleField: 'Title'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data.parameters, {
      Title: validTitle,
      ViewFields: {
        results: ['StartDate', 'EndDate', 'Title']
      },
      CalendarViewStyles: '<CalendarViewStyle Title="Day" Type="day" Template="CalendarViewdayChrome" Sequence="1" Default="FALSE" /><CalendarViewStyle Title="Week" Type="week" Template="CalendarViewweekChrome" Sequence="2" Default="FALSE" /><CalendarViewStyle Title="Month" Type="month" Template="CalendarViewmonthChrome" Sequence="3" Default="TRUE" /><CalendarViewStyle Title="Work week" Type="workweek" Template="CalendarViewweekChrome" Sequence="4" Default="FALSE" />',
      Query: `<Where><DateRangesOverlap><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><Value Type='DateTime'><Month /></Value></DateRangesOverlap></Where>`,
      ViewData: '<FieldRef Name="Title" Type="CalendarMonthTitle" /><FieldRef Name="Title" Type="CalendarWeekTitle" /><FieldRef Name="" Type="CalendarWeekLocation" /><FieldRef Name="Title" Type="CalendarDayTitle" /><FieldRef Name="" Type="CalendarDayLocation" />',
      CustomFormatter: undefined,
      PersonalView: false,
      SetAsDefaultView: false,
      Paged: true,
      RowLimit: 30,
      ViewType2: 'MODERNCALENDAR'
    });
  });

  it('correctly adds a calendar view with additional options', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(validListId)}')/views/add`) {
        return viewCreationResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listId: validListId,
        type: 'calendar',
        title: validTitle,
        fields: validFieldsInput,
        calendarStartDateField: 'StartDate',
        calendarEndDateField: 'EndDate',
        calendarTitleField: 'Title',
        calendarSubTitleField: 'Subtitle',
        calendarDefaultLayout: 'workWeek',
        customFormatter: JSON.stringify({ someProperty: 'someValue' })
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data.parameters, {
      Title: validTitle,
      ViewFields: {
        results: ['StartDate', 'EndDate', 'Title', 'Subtitle', ...validFieldsInput.split(',')]
      },
      CalendarViewStyles: '<CalendarViewStyle Title="Day" Type="day" Template="CalendarViewdayChrome" Sequence="1" Default="FALSE" /><CalendarViewStyle Title="Week" Type="week" Template="CalendarViewweekChrome" Sequence="2" Default="FALSE" /><CalendarViewStyle Title="Month" Type="month" Template="CalendarViewmonthChrome" Sequence="3" Default="FALSE" /><CalendarViewStyle Title="Work week" Type="workweek" Template="CalendarViewweekChrome" Sequence="4" Default="TRUE" />',
      Query: `<Where><DateRangesOverlap><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><Value Type='DateTime'><Month /></Value></DateRangesOverlap></Where>`,
      ViewData: '<FieldRef Name="Title" Type="CalendarMonthTitle" /><FieldRef Name="Title" Type="CalendarWeekTitle" /><FieldRef Name="Subtitle" Type="CalendarWeekLocation" /><FieldRef Name="Title" Type="CalendarDayTitle" /><FieldRef Name="Subtitle" Type="CalendarDayLocation" />',
      CustomFormatter: JSON.stringify({ someProperty: 'someValue' }),
      PersonalView: false,
      SetAsDefaultView: false,
      Paged: true,
      RowLimit: 30,
      ViewType2: 'MODERNCALENDAR'
    });
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listUrl: validListUrl,
        title: validTitle,
        fields: validFieldsInput,
        rowLimit: 100
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
