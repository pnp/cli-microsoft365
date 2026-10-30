import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import command, { options } from './file-roleassignment-list.js';

describe(commands.FILE_ROLEASSIGNMENT_LIST, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const fileId = 'b2307a39-e878-4586-8901-08ff728c5496';
  const fileUrl = '/sites/project-x/Documents/Test1.docx';
  const fileResponse = {
    CheckInComment: '',
    CheckOutType: 2,
    ContentTag: '{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12',
    CustomizedPageStatus: 0,
    ETag: '"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9"',
    Exists: true,
    IrmEnabled: false,
    Length: '331673',
    Level: 1,
    LinkingUri: 'https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    LinkingUrl: 'https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    ListItemAllFields: {
      Id: 1,
      ID: 1
    },
    MajorVersion: 3,
    MinorVersion: 0,
    Name: 'Test1.docx',
    ServerRelativeUrl: '/sites/project-x/documents/Test1.docx',
    TimeCreated: '2018-02-05T08:42:36Z',
    TimeLastModified: '2018-02-05T08:44:03Z',
    Title: '',
    UIVersion: 1536,
    UIVersionLabel: '3.0',
    UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
  };
  const fileRoleAssignmentsResponse = {
    value: [{
      Member: {
        Id: 3,
        IsHiddenInUI: false,
        LoginName: "Communication site Owners",
        Title: "Communication site Owners",
        PrincipalType: 8,
        AllowMembersEditMembership: false,
        AllowRequestToJoinLeave: false,
        AutoAcceptRequestToJoinLeave: false,
        Description: null,
        OnlyAllowMembersViewMembership: false,
        OwnerTitle: "Communication site Owners",
        RequestToJoinLeaveEmailSetting: ""
      },
      RoleDefinitionBindings: [
        {
          BasePermissions: {
            High: "2147483647",
            Low: "4294967295"
          },
          Description: "Has full control.",
          Hidden: false,
          Id: 1073741829,
          Name: "Full Control",
          Order: 1,
          RoleTypeKind: 5
        }
      ],
      PrincipalId: 32
    }]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let commandToExecuteStub: sinon.SinonStub | undefined;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      spo.getFileById,
      spo.getFileByUrl,
      spo.getFileRoleAssignments
    ]);
    commandToExecuteStub?.restore();
    commandToExecuteStub = undefined;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_ROLEASSIGNMENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('has the correct default properties', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['PrincipalId', 'Member', 'RoleDefinitionBindings']);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', fileId: fileId });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if the fileId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileId: 'foo' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo is not a valid GUID');
  });

  it('passes validation if webUrl and fileId are valid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileId: fileId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if webUrl and fileUrl are valid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileUrl: fileUrl });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if neither fileId nor fileUrl is specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both fileId and fileUrl are specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileId: fileId, fileUrl: fileUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileUrl: fileUrl, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('tracks whether the fileId and fileUrl options are specified in telemetry', () => {
    const getTelemetryProperties = (command as any).getTelemetryProperties.bind(command);
    assert.deepStrictEqual(getTelemetryProperties({ options: { fileId: fileId } }), { fileId: true });
    assert.deepStrictEqual(getTelemetryProperties({ options: { fileUrl: fileUrl } }), { fileUrl: true });
  });

  it('retrieves file role assignments by file id', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').resolves(fileRoleAssignmentsResponse.value);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: webUrl, fileId: fileId }) });
    assert(loggerLogSpy.calledWith(fileRoleAssignmentsResponse.value));
  });

  it('retrieves file role assignments by file url', async () => {
    sinon.stub(spo, 'getFileByUrl').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').resolves(fileRoleAssignmentsResponse.value);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: webUrl, fileUrl: fileUrl }) });
    assert(loggerLogSpy.calledWith(fileRoleAssignmentsResponse.value));
  });

  it('preserves nested role assignments for JSON output', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').resolves(fileRoleAssignmentsResponse.value);
    const options = commandOptionsSchema.parse({ webUrl, fileId, output: 'json' });

    await command.action(logger, { options });

    assert(loggerLogSpy.calledWith(fileRoleAssignmentsResponse.value));
    assert.deepStrictEqual(JSON.parse(await cli.formatOutput(command, loggerLogSpy.firstCall.args[0], options)), fileRoleAssignmentsResponse.value);
  });

  ['text', 'csv', 'md'].forEach(output => {
    it(`renders member titles and comma-separated role names in ${output} output`, async () => {
      const assignments = [
        {
          ...fileRoleAssignmentsResponse.value[0],
          RoleDefinitionBindings: [
            ...fileRoleAssignmentsResponse.value[0].RoleDefinitionBindings,
            { Name: 'Limited Access' }
          ]
        },
        {
          PrincipalId: 33,
          Member: { Title: 'Communication site Visitors' },
          RoleDefinitionBindings: [{ Name: 'Read' }]
        }
      ];
      const originalAssignments = structuredClone(assignments);
      sinon.stub(spo, 'getFileByUrl').resolves(fileResponse);
      sinon.stub(spo, 'getFileRoleAssignments').resolves(assignments);
      commandToExecuteStub = sinon.stub(cli, 'commandToExecute').value(commandInfo);
      const options = commandOptionsSchema.parse({ webUrl, fileUrl, output });

      await command.action(logger, { options });

      const result = loggerLogSpy.firstCall.args[0];
      assert.deepStrictEqual(result, [
        {
          PrincipalId: 32,
          Member: 'Communication site Owners',
          RoleDefinitionBindings: 'Full Control, Limited Access'
        },
        {
          PrincipalId: 33,
          Member: 'Communication site Visitors',
          RoleDefinitionBindings: 'Read'
        }
      ]);
      const formatted = await cli.formatOutput(command, result, options);
      assert(!formatted.includes('[object Object]'));
      if (output === 'text') {
        const rows = formatted.trim().split(/\r?\n/)
          .filter((row: string) => !row.startsWith('---'))
          .map((row: string) => row.trim().replace(/\s{2,}/g, '|'));
        assert.deepStrictEqual(rows, [
          'PrincipalId|Member|RoleDefinitionBindings',
          '32|Communication site Owners|Full Control, Limited Access',
          '33|Communication site Visitors|Read'
        ]);
      }
      else if (output === 'csv') {
        assert.strictEqual(formatted, 'PrincipalId,Member,RoleDefinitionBindings\n32,Communication site Owners,"Full Control, Limited Access"\n33,Communication site Visitors,Read\n');
      }
      else {
        assert(formatted.includes('PrincipalId | 32'));
        assert(formatted.includes('Member | Communication site Owners'));
        assert(formatted.includes('RoleDefinitionBindings | Full Control, Limited Access'));
        assert(formatted.includes('PrincipalId | 33'));
        assert(formatted.includes('Member | Communication site Visitors'));
        assert(formatted.includes('RoleDefinitionBindings | Read'));
      }
      assert.deepStrictEqual(assignments, originalAssignments);
    });

    it(`handles no role assignments in ${output} output`, async () => {
      sinon.stub(spo, 'getFileById').resolves(fileResponse);
      sinon.stub(spo, 'getFileRoleAssignments').resolves([]);

      await command.action(logger, { options: commandOptionsSchema.parse({ webUrl, fileId, output }) });

      assert(loggerLogSpy.calledOnceWithExactly([]));
    });
  });

  it('handles an assignment with no role definition bindings', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').resolves([{
      ...fileRoleAssignmentsResponse.value[0],
      RoleDefinitionBindings: []
    }]);

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl, fileId, output: 'text' }) });

    assert(loggerLogSpy.calledOnceWithExactly([{
      PrincipalId: 32,
      Member: 'Communication site Owners',
      RoleDefinitionBindings: ''
    }]));
  });

  it('correctly handles error when retrieving file role assignments', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: webUrl, fileId: fileId }) }), new CommandError('An error has occurred'));
  });
});