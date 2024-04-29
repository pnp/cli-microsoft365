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
import command from './file-roleassignment-add.js';
import { settingsNames } from '../../../../settingsNames.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.FILE_ROLEASSIGNMENT_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const fileUrl = '/sites/project-x/documents/Test1.docx';
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const roleDefinitionResponse = {
    BasePermissions: {
      High: 2147483647,
      Low: 4294967295
    },
    Description: 'Has full control.',
    Hidden: false,
    Id: 1073741827,
    Name: 'Full Control',
    Order: 1,
    RoleTypeKind: 5,
    BasePermissionsValue: [
      'ViewListItems',
      'AddListItems',
      'EditListItems',
      'DeleteListItems',
      'ApproveItems',
      'OpenItems',
      'ViewVersions',
      'DeleteVersions',
      'CancelCheckout',
      'ManagePersonalViews',
      'ManageLists',
      'ViewFormPages',
      'AnonymousSearchAccessList',
      'Open',
      'ViewPages',
      'AddAndCustomizePages',
      'ApplyThemeAndBorder',
      'ApplyStyleSheets',
      'ViewUsageData',
      'CreateSSCSite',
      'ManageSubwebs',
      'CreateGroups',
      'ManagePermissions',
      'BrowseDirectories',
      'BrowseUserInfo',
      'AddDelPrivateWebParts',
      'UpdatePersonalWebParts',
      'ManageWeb',
      'AnonymousSearchAccessWebLists',
      'UseClientIntegration',
      'UseRemoteAPIs',
      'ManageAlerts',
      'CreateAlerts',
      'EditMyUserInfo',
      'EnumeratePermissions'
    ],
    RoleTypeKindValue: 'Administrator'
  };
  const fileResponse = {
    CheckInComment: '',
    CheckOutType: 2,
    ContentTag: '{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12',
    CustomizedPageStatus: 0,
    ETag: '\'{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\'',
    Exists: true,
    IrmEnabled: false,
    Length: '331673',
    Level: 1,
    LinkingUri: 'https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    LinkingUrl: 'https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    MajorVersion: 3,
    MinorVersion: 0,
    Name: 'Test1.docx',
    ServerRelativeUrl: '/sites/project-x/documents/Test1.docx',
    TimeCreated: '2018-02-05T08:42:36Z',
    TimeLastModified: '2018-02-05T08:44:03Z',
    Title: '',
    UIVersion: 1536,
    UIVersionLabel: '3.0',
    UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6',
    ListItemAllFields: {
      Id: 4,
      ID: 4
    }
  };
  const userResponse = {
    Id: 11,
    IsHiddenInUI: false,
    LoginName: 'i:0#.f|membership|someaccount@tenant.onmicrosoft.com',
    Title: 'Some Account',
    PrincipalType: 1,
    Email: 'someaccount@tenant.onmicrosoft.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: true,
    UserId: {
      NameId: '1003200097d06dd6',
      NameIdIssuer: 'urn:federation:microsoftonline'
    },
    UserPrincipalName: 'someaccount@tenant.onmicrosoft.com'
  };
  const groupResponse = {
    Id: 5,
    IsHiddenInUI: false,
    LoginName: "Group A",
    Title: "Group A",
    PrincipalType: 8,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    Description: "",
    OnlyAllowMembersViewMembership: true,
    OwnerTitle: "Some Account",
    RequestToJoinLeaveEmailSetting: null
  };

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      spo.getRoleDefinitionByName,
      spo.getGroupByName,
      spo.getUserByEmail,
      spo.getFileById,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, groupName: 'Group name A', roleDefinitionName: 'Read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo', groupName: 'Group name A', roleDefinitionName: 'Read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, principalId: 'NaN', roleDefinitionName: 'Read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the roleDefinitionId option is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, groupName: 'Group name A', roleDefinitionId: 'NaN' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if no roledefinition is passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, principalId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, groupName: 'Group name A', roleDefinitionName: 'Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles error when adding file role assignment', async () => {
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
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: 10,
        roleDefinitionId: 1073741827
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('correctly adds role assignment specifying principalId and role definition name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='10',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getRoleDefinitionByName').resolves(roleDefinitionResponse);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: 10,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly adds role assignment specifying principalId and role definition name, retrieving file by the ID', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='10',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getFileById').resolves(fileResponse);
    sinon.stub(spo, 'getRoleDefinitionByName').resolves(roleDefinitionResponse);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId,
        principalId: 10,
        roleDefinitionName: 'Full Control'
      }
    });
  });

  it('correctly adds role assignment specifying upn and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='11',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getUserByEmail').resolves(userResponse);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    const error = 'no user found';
    sinon.stub(spo, 'getUserByEmail').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    }), new CommandError('no user found'));
  });

  it('correctly adds role assignment specifying groupName and role definition id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url as string === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2Fdocuments%2FTest1.docx')/ListItemAllFields/roleassignments/addroleassignment(principalid='5',roledefid='1073741827')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getGroupByName').resolves(groupResponse);

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        groupName: 'Group A',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when role definition does not exist', async () => {
    const error = 'no role definition found';
    sinon.stub(spo, 'getRoleDefinitionByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        groupName: 'Group A',
        roleDefinitionName: 'Non-existing Role Definition'
      }
    }), new CommandError('no role definition found'));
  });

  it('correctly handles error when group does not exist', async () => {
    const error = 'no group found';
    sinon.stub(spo, 'getGroupByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        fileUrl: fileUrl,
        groupName: 'Group A',
        roleDefinitionId: 1073741827
      }
    }), new CommandError('no group found'));
  });
});
