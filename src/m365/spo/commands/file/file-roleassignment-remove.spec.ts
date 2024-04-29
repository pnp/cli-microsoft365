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
import commands from '../../commands.js';
import command from './file-roleassignment-remove.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.FILE_ROLEASSIGNMENT_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/contoso-sales';
  const fileUrl = '/sites/contoso-sales/documents/Test1.docx';
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const principalId = 2;
  const upn = 'user1@contoso.onmicrosoft.com';
  const groupName = 'saleGroup';
  const fileResponse = {
    CheckInComment: '',
    CheckOutType: 2,
    ContentTag: '{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12',
    CustomizedPageStatus: 0,
    ETag: '\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"',
    Exists: true,
    IrmEnabled: false,
    Length: '331673',
    Level: 1,
    LinkingUri: 'https://contoso.sharepoint.com/sites/contoso-sales/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    LinkingUrl: 'https://contoso.sharepoint.com/sites/contoso-sales/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    ListItemAllFields: {
      Id: 1,
      ID: 1
    },
    MajorVersion: 3,
    MinorVersion: 0,
    Name: 'Test1.docx',
    ServerRelativeUrl: '/sites/contoso-sales/documents/Test1.docx',
    TimeCreated: '2018-02-05T08:42:36Z',
    TimeLastModified: '2018-02-05T08:44:03Z',
    Title: '',
    UIVersion: 1536,
    UIVersionLabel: '3.0',
    UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
  };
  const userResponse = {
    Id: 2,
    IsHiddenInUI: false,
    LoginName: 'i:0#.f|membership|user1@contoso.onmicrosoft.com',
    Title: 'User1',
    PrincipalType: 1,
    Email: 'user1@contoso.onmicrosoft.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: {
      NameId: '10032002473c5ae3',
      NameIdIssuer: 'urn:federation:microsoftonline'
    },
    UserPrincipalName: 'user1@contoso.onmicrosoft.com'
  };

  const groupResponse = {
    Id: 2,
    IsHiddenInUI: false,
    LoginName: "saleGroup",
    Title: "saleGroup",
    PrincipalType: 8,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    Description: "",
    OnlyAllowMembersViewMembership: true,
    OwnerTitle: "John Doe",
    RequestToJoinLeaveEmailSetting: null
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      cli.promptForConfirmation,
      spo.getUserByEmail,
      spo.getGroupByName,
      spo.getFileById,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_ROLEASSIGNMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, principalId: principalId, force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo', principalId: principalId, force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, principalId: 'Hi', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '0cd891ef-afce-4e55-b836-fce03286cccf', principalId: principalId, force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing role assignment from the file when force option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId,
        principalId: principalId
      }
    });


    assert(promptIssued);
  });

  it('aborts removing role assignment from the file when force option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId,
        principalId: principalId
      }
    });

    assert(postSpy.notCalled);
  });

  it('remove role assignment from the file by relative URL and principal Id when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: principalId
      }
    });
  });

  it('remove role assignment from the file by Id and upn', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getUserByEmail').resolves(userResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileId: fileId,
        upn: upn,
        force: true
      }
    });
  });

  it('remove role assignment from the file by Id and group name', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getGroupByName').resolves(groupResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileId: fileId,
        groupName: groupName,
        force: true
      }
    });
  });

  it('correctly handles error when removing file role assignment', async () => {
    const errorMessage = 'request rejected';
    sinon.stub(request, 'post').callsFake(async () => { throw errorMessage; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: principalId,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
