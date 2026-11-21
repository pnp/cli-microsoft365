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
import command, { options } from './web-roleassignment-list.js';

describe(commands.WEB_ROLEASSIGNMENT_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const webUrl = 'https://contoso.sharepoint.com/sites/Marketing';

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
            BasePermissions: { High: 2147483647, Low: 4294967295 },
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
          Id: 5,
          IsHiddenInUI: false,
          LoginName: 'Marketing Members',
          Title: 'Marketing Members',
          PrincipalType: 8,
          AllowMembersEditMembership: true,
          AllowRequestToJoinLeave: false,
          AutoAcceptRequestToJoinLeave: false,
          Description: null,
          OnlyAllowMembersViewMembership: false,
          OwnerTitle: 'Marketing Owners',
          RequestToJoinLeaveEmailSetting: ''
        },
        RoleDefinitionBindings: [
          {
            BasePermissions: { High: 432, Low: 1011028719 },
            Description: 'Can view, add, update, delete, approve, and customize.',
            Hidden: false,
            Id: 1073741827,
            Name: 'Edit',
            Order: 48,
            RoleTypeKind: 6
          }
        ],
        PrincipalId: 5
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.schema as typeof options;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => { log.push(msg); },
      logRaw: async (msg: string) => { log.push(msg); },
      logToStderr: async (msg: string) => { log.push(msg); }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([request.get]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.WEB_ROLEASSIGNMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct default properties', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['PrincipalId', 'Member', 'RoleDefinitionBindings']);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'invalid' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if webUrl is a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves role assignments for a web (JSON output)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/RoleAssignments?$expand=Member,RoleDefinitionBindings`) {
        return roleAssignmentsResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl, output: 'json' } });
    assert(loggerLogSpy.calledOnce);
    const result = loggerLogSpy.firstCall.args[0];
    assert.strictEqual(result.length, 2);
    assert.strictEqual(result[0].PrincipalId, 3);
  });

  it('retrieves role assignments for a web (text output, flattens Member and RoleDefinitionBindings)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/RoleAssignments?$expand=Member,RoleDefinitionBindings`) {
        return roleAssignmentsResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl, output: 'text' } });
    assert(loggerLogSpy.calledOnce);
    const result = loggerLogSpy.firstCall.args[0];
    assert.strictEqual(result[0].Member, 'Marketing Owners');
    assert.strictEqual(typeof result[0].RoleDefinitionBindings, 'string');
  });

  it('outputs verbose logging', async () => {
    sinon.stub(request, 'get').resolves(roleAssignmentsResponse);
    const logToStderrSpy = sinon.spy(logger, 'logToStderr');

    await command.action(logger, { options: { webUrl, verbose: true } });
    assert(logToStderrSpy.calledWith(sinon.match(/Retrieving role assignments/)));
  });

  it('handles error from API', async () => {
    sinon.stub(request, 'get').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(
      command.action(logger, { options: { webUrl } }),
      (err: CommandError) => err.message === 'An error has occurred'
    );
  });
});
