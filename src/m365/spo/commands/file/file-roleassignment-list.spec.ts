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
import command from './file-roleassignment-list.js';

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

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if webUrl and fileUrl are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileUrl: fileUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves file role assignments by file id', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').resolves(fileRoleAssignmentsResponse.value);

    await command.action(logger, { options: { debug: true, webUrl: webUrl, fileId: fileId } });
    assert(loggerLogSpy.calledWith(fileRoleAssignmentsResponse.value));
  });

  it('retrieves file role assignments by file url', async () => {
    sinon.stub(spo, 'getFileByUrl').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').resolves(fileRoleAssignmentsResponse.value);

    await command.action(logger, { options: { debug: true, webUrl: webUrl, fileUrl: fileUrl } });
    assert(loggerLogSpy.calledWith(fileRoleAssignmentsResponse.value));
  });

  it('correctly handles error when retrieving file role assignments', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getFileRoleAssignments').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: webUrl, fileId: fileId } }), new CommandError('An error has occurred'));
  });
});