import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './listitem-roleassignment-list.js';

describe(commands.LISTITEM_ROLEASSIGNMENT_LIST, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const listId = '0cd891ef-afce-4e55-b836-fce03286cccf';
  const listTitle = 'Branding';
  const listItemId = 125;
  const roleAssignmentsPath = `/items(${listItemId})/RoleAssignments?$expand=Member,RoleDefinitionBindings`;

  const roleAssignmentsResponse = {
    value: [
      {
        Member: {
          Id: 3,
          IsHiddenInUI: false,
          LoginName: 'Marketing Owners',
          Title: 'Marketing Owners',
          PrincipalType: 8,
          AllowMembersEditMembership: false,
          AllowRequestToJoinLeave: false,
          AutoAcceptRequestToJoinLeave: false,
          Description: null,
          OnlyAllowMembersViewMembership: false,
          OwnerTitle: 'Marketing Owners',
          RequestToJoinLeaveEmailSetting: ''
        },
        RoleDefinitionBindings: [
          {
            BasePermissions: {
              High: '2147483647',
              Low: '4294967295'
            },
            Description: 'Has full control.',
            Hidden: false,
            Id: 1073741829,
            Name: 'Full Control',
            Order: 1,
            RoleTypeKind: 5
          }
        ],
        PrincipalId: 3
      },
      {
        Member: {
          Id: 11,
          IsHiddenInUI: false,
          LoginName: 'i:0#.f|membership|john@contoso.onmicrosoft.com',
          Title: 'John Doe',
          PrincipalType: 1,
          Email: 'john@contoso.onmicrosoft.com',
          Expiration: '',
          IsEmailAuthenticationGuestUser: false,
          IsShareByEmailGuestUser: false,
          IsSiteAdmin: false,
          UserId: {
            NameId: '1003200097d06dd6',
            NameIdIssuer: 'urn:federation:microsoftonline'
          },
          UserPrincipalName: 'john@contoso.onmicrosoft.com'
        },
        RoleDefinitionBindings: [
          {
            BasePermissions: {
              High: '432',
              Low: '1012866047'
            },
            Description: 'Can view, add, update, delete, approve, and customize.',
            Hidden: false,
            Id: 1073741830,
            Name: 'Design',
            Order: 32,
            RoleTypeKind: 4
          },
          {
            BasePermissions: {
              High: '176',
              Low: '138612833'
            },
            Description: 'Can view pages and list items and download documents.',
            Hidden: false,
            Id: 1073741826,
            Name: 'Read',
            Order: 128,
            RoleTypeKind: 2
          }
        ],
        PrincipalId: 11
      }
    ]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ROLEASSIGNMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct default properties', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['PrincipalId', 'Member', 'RoleDefinitionBindings']);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listId: listId, listItemId: listItemId });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: listItemId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: '12345', listItemId: listItemId });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, '12345 is not a valid GUID');
  });

  it('fails validation if the listItemId option is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the listItemId option is not a positive integer', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: -1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId, listTitle and listUrl are not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listItemId: listItemId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both listId and listTitle are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listTitle: listTitle, listItemId: listItemId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listTitle is an empty string', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listTitle: '', listItemId: listItemId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listUrl is an empty string', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listUrl: '', listItemId: listItemId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: listItemId, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('lists role assignments of a list item in a list specified by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/getByTitle('${listTitle}')${roleAssignmentsPath}`) {
        return structuredClone(roleAssignmentsResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId } });
    const result = loggerLogSpy.lastCall.args[0];
    assert.strictEqual(result.length, 2);
    assert.strictEqual(result[0].PrincipalId, 3);
    assert.strictEqual(result[0].Member.Title, 'Marketing Owners');
    assert.strictEqual(result[0].RoleDefinitionBindings[0].RoleTypeKindValue, 'Administrator');
    assert(result[0].RoleDefinitionBindings[0].BasePermissionsValue.includes('ManagePermissions'));
  });

  it('lists role assignments of a list item in a list specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'${listId}')${roleAssignmentsPath}`) {
        return structuredClone(roleAssignmentsResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listId: listId, listItemId: listItemId, output: 'json' } });
    const result = loggerLogSpy.lastCall.args[0];
    assert.strictEqual(result.length, 2);
    assert.strictEqual(result[1].Member.Title, 'John Doe');
  });

  it('lists role assignments of a list item in a list specified by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('%2Fsites%2FMarketing%2FBranding')${roleAssignmentsPath}`) {
        return structuredClone(roleAssignmentsResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listUrl: '/Branding', listItemId: listItemId, verbose: true } });
    assert(loggerLogToStderrSpy.calledWith(`Retrieving role assignments of list item with ID ${listItemId}...`));
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 2);
  });

  it('lists role assignments of a list item with member title and comma-separated role names in non-JSON output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/getByTitle('${listTitle}')${roleAssignmentsPath}`) {
        return structuredClone(roleAssignmentsResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, output: 'csv' } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        PrincipalId: 3,
        Member: 'Marketing Owners',
        RoleDefinitionBindings: 'Full Control'
      },
      {
        PrincipalId: 11,
        Member: 'John Doe',
        RoleDefinitionBindings: 'Design, Read'
      }
    ]));
  });

  it('logs the full role assignments in non-JSON output when a query is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/getByTitle('${listTitle}')${roleAssignmentsPath}`) {
        return structuredClone(roleAssignmentsResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, output: 'text', query: '[].Member.Title' } });
    const logged = loggerLogSpy.lastCall.args[0];
    assert.strictEqual(logged[0].Member.Title, 'Marketing Owners');
    assert.strictEqual(logged[1].RoleDefinitionBindings[0].Name, 'Design');
  });

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-2147024809, System.ArgumentException',
          message: {
            value: 'Item does not exist. It may have been deleted by another user.'
          }
        }
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId } }),
      new CommandError(error.error['odata.error'].message.value));
  });
});
