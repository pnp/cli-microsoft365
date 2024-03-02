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
import command from './file-sharinginfo-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.FILE_SHARINGINFO_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const JSONOuput = {
    "permissionsInformation": {
      "hasInheritedLinks": false,
      "links": [{
        "isInherited": false,
        "linkDetails": {
          "AllowsAnonymousAccess": false,
          "ApplicationId": null,
          "BlocksDownload": false,
          "Created": "2020-11-03T07:49:04.928Z",
          "CreatedBy": {
            "email": "user1@contoso.onmicrosoft.com",
            "expiration": "",
            "id": 6,
            "isActive": true,
            "isExternal": false,
            "jobTitle": "CEO",
            "loginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com",
            "name": "User 1",
            "principalType": 0,
            "userId": null,
            "userPrincipalName": "user1@contoso.onmicrosoft.com"
          },
          "Description": null,
          "Embeddable": false,
          "Expiration": "",
          "HasExternalGuestInvitees": true,
          "Invitations": [{
            "invitedBy": {
              "email": "user1@contoso.onmicrosoft.com",
              "expiration": "",
              "id": 6,
              "isActive": true,
              "isExternal": false,
              "jobTitle": "CEO",
              "loginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com",
              "name": "User 1",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "user1@contoso.onmicrosoft.com"
            },
            "invitedOn": "2020-11-03T07:49:04.803Z",
            "invitee": {
              "email": "user2@contoso.onmicrosoft.com",
              "expiration": "",
              "id": 17,
              "isActive": true,
              "isExternal": false,
              "jobTitle": "CEO",
              "loginName": "i:0#.f|membership|user2@contoso.onmicrosoft.com",
              "name": "User 2",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "user2@contoso.onmicrosoft.com"
            }
          }, {
            "invitedBy": {
              "email": "user1@contoso.onmicrosoft.com",
              "expiration": "",
              "id": 6,
              "isActive": true,
              "isExternal": false,
              "jobTitle": "CEO",
              "loginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com",
              "name": "User 1",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "user1@contoso.onmicrosoft.com"
            },
            "invitedOn": "2020-11-03T07:49:04.803Z",
            "invitee": {
              "email": "user3@contoso.onmicrosoft.com",
              "expiration": "",
              "id": 18,
              "isActive": true,
              "isExternal": false,
              "jobTitle": "Executive Vice President",
              "loginName": "i:0#.f|membership|user3@contoso.onmicrosoft.com",
              "name": "User 3",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "user3@contoso.onmicrosoft.com"
            }
          }, {
            "invitedBy": {
              "email": "user1@contoso.onmicrosoft.com",
              "expiration": "",
              "id": 6,
              "isActive": true,
              "isExternal": false,
              "jobTitle": "CEO",
              "loginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com",
              "name": "User 1",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "user1@contoso.onmicrosoft.com"
            },
            "invitedOn": "2020-11-03T07:49:04.803Z",
            "invitee": {
              "email": "external1@externalcontoso.onmicrosoft.com",
              "expiration": "",
              "id": 23,
              "isActive": true,
              "isExternal": true,
              "jobTitle": null,
              "loginName": "i:0#.f|membership|external1_externalcontoso#ext#@contoso.onmicrosoft.com",
              "name": "External User 1",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "external1_externalcontoso#ext#@contoso.onmicrosoft.com"
            }
          }, {
            "invitedBy": {
              "email": "user1@contoso.onmicrosoft.com",
              "expiration": "",
              "id": 6,
              "isActive": true,
              "isExternal": false,
              "jobTitle": "CEO",
              "loginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com",
              "name": "User 1",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "user1@contoso.onmicrosoft.com"
            },
            "invitedOn": "2020-11-05T11:10:20.239Z",
            "invitee": {
              "email": "user4@contoso.onmicrosoft.com",
              "expiration": "",
              "id": 20,
              "isActive": true,
              "isExternal": false,
              "jobTitle": "Executive Vice President",
              "loginName": "i:0#.f|membership|user4@contoso.onmicrosoft.com",
              "name": "User 4",
              "principalType": 1,
              "userId": null,
              "userPrincipalName": "user4@contoso.onmicrosoft.com"
            }
          }],
          "IsActive": true,
          "IsAddressBarLink": false,
          "IsCreateOnlyLink": false,
          "IsDefault": true,
          "IsEditLink": false,
          "IsFormsLink": false,
          "IsReviewLink": false,
          "IsUnhealthy": false,
          "LastModified": "2020-11-05T12:11:05.914Z",
          "LastModifiedBy": {
            "email": "user1@contoso.onmicrosoft.com",
            "expiration": "",
            "id": 6,
            "isActive": true,
            "isExternal": false,
            "jobTitle": "CEO",
            "loginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com",
            "name": "User 1",
            "principalType": 0,
            "userId": null,
            "userPrincipalName": "user1@contoso.onmicrosoft.com"
          },
          "LimitUseToApplication": false,
          "LinkKind": 6,
          "PasswordLastModified": "",
          "PasswordLastModifiedBy": null,
          "RedeemedUsers": [],
          "RequiresPassword": false,
          "RestrictedShareMembership": true,
          "Scope": 2,
          "ShareId": "6205a23e-e1fb-49e0-ba6d-e4ae76638de3",
          "ShareTokenString": "share=EVv36NDX3I5HnKpBlE8h1yIBLsx9EUjgBSgEN0fnh-3tQA",
          "SharingLinkStatus": 0,
          "TrackLinkUsers": false,
          "Url": "https://contoso.sharepoint.com/:w:/s/project-x/EVv36NDX3I5HnKpBlE8h1yIBLsx9EUjgBSgEN0fnh-3tQA"
        },
        "linkMembers": [{
          "email": "external1@externalcontoso.onmicrosoft.com",
          "expiration": "",
          "id": 23,
          "isActive": true,
          "isExternal": true,
          "jobTitle": null,
          "loginName": "i:0#.f|membership|external1_externalcontoso#ext#@contoso.onmicrosoft.com",
          "name": "External User 1",
          "principalType": 1,
          "userId": null,
          "userPrincipalName": "external1_externalcontoso#ext#@contoso.onmicrosoft.com"
        }, {
          "email": "user2@contoso.onmicrosoft.com",
          "expiration": "",
          "id": 17,
          "isActive": true,
          "isExternal": false,
          "jobTitle": "CEO",
          "loginName": "i:0#.f|membership|user2@contoso.onmicrosoft.com",
          "name": "User 2",
          "principalType": 1,
          "userId": null,
          "userPrincipalName": "user2@contoso.onmicrosoft.com"
        }, {
          "email": "user4@contoso.onmicrosoft.com",
          "expiration": "",
          "id": 20,
          "isActive": true,
          "isExternal": false,
          "jobTitle": "Executive Vice President",
          "loginName": "i:0#.f|membership|user4@contoso.onmicrosoft.com",
          "name": "User 4",
          "principalType": 1,
          "userId": null,
          "userPrincipalName": "user4@contoso.onmicrosoft.com"
        }, {
          "email": "user3@contoso.onmicrosoft.com",
          "expiration": "",
          "id": 18,
          "isActive": true,
          "isExternal": false,
          "jobTitle": "Executive Vice President",
          "loginName": "i:0#.f|membership|user3@contoso.onmicrosoft.com",
          "name": "User 3",
          "principalType": 1,
          "userId": null,
          "userPrincipalName": "user3@contoso.onmicrosoft.com"
        }]
      }, {
        "isInherited": false,
        "linkDetails": {
          "AllowsAnonymousAccess": false,
          "ApplicationId": null,
          "BlocksDownload": false,
          "Created": "",
          "CreatedBy": null,
          "Description": null,
          "Embeddable": false,
          "Expiration": "",
          "HasExternalGuestInvitees": false,
          "Invitations": [],
          "IsActive": false,
          "IsAddressBarLink": false,
          "IsCreateOnlyLink": false,
          "IsDefault": true,
          "IsEditLink": false,
          "IsFormsLink": false,
          "IsReviewLink": false,
          "IsUnhealthy": false,
          "LastModified": "",
          "LastModifiedBy": null,
          "LimitUseToApplication": false,
          "LinkKind": 5,
          "PasswordLastModified": "",
          "PasswordLastModifiedBy": null,
          "RedeemedUsers": [],
          "RequiresPassword": false,
          "RestrictedShareMembership": false,
          "Scope": -1,
          "ShareId": "00000000-0000-0000-0000-000000000000",
          "ShareTokenString": null,
          "SharingLinkStatus": 0,
          "TrackLinkUsers": false,
          "Url": null
        },
        "linkMembers": []
      }, {
        "isInherited": false,
        "linkDetails": {
          "AllowsAnonymousAccess": false,
          "ApplicationId": null,
          "BlocksDownload": false,
          "Created": "",
          "CreatedBy": null,
          "Description": null,
          "Embeddable": false,
          "Expiration": "",
          "HasExternalGuestInvitees": false,
          "Invitations": [],
          "IsActive": false,
          "IsAddressBarLink": false,
          "IsCreateOnlyLink": false,
          "IsDefault": true,
          "IsEditLink": false,
          "IsFormsLink": false,
          "IsReviewLink": false,
          "IsUnhealthy": false,
          "LastModified": "",
          "LastModifiedBy": null,
          "LimitUseToApplication": false,
          "LinkKind": 4,
          "PasswordLastModified": "",
          "PasswordLastModifiedBy": null,
          "RedeemedUsers": [],
          "RequiresPassword": false,
          "RestrictedShareMembership": false,
          "Scope": -1,
          "ShareId": "00000000-0000-0000-0000-000000000000",
          "ShareTokenString": null,
          "SharingLinkStatus": 0,
          "TrackLinkUsers": false,
          "Url": null
        },
        "linkMembers": []
      }, {
        "isInherited": false,
        "linkDetails": {
          "AllowsAnonymousAccess": false,
          "ApplicationId": null,
          "BlocksDownload": false,
          "Created": "",
          "CreatedBy": null,
          "Description": null,
          "Embeddable": false,
          "Expiration": "",
          "HasExternalGuestInvitees": false,
          "Invitations": [],
          "IsActive": false,
          "IsAddressBarLink": false,
          "IsCreateOnlyLink": false,
          "IsDefault": true,
          "IsEditLink": false,
          "IsFormsLink": false,
          "IsReviewLink": false,
          "IsUnhealthy": false,
          "LastModified": "",
          "LastModifiedBy": null,
          "LimitUseToApplication": false,
          "LinkKind": 3,
          "PasswordLastModified": "",
          "PasswordLastModifiedBy": null,
          "RedeemedUsers": [],
          "RequiresPassword": false,
          "RestrictedShareMembership": false,
          "Scope": -1,
          "ShareId": "00000000-0000-0000-0000-000000000000",
          "ShareTokenString": null,
          "SharingLinkStatus": 0,
          "TrackLinkUsers": false,
          "Url": null
        },
        "linkMembers": []
      }, {
        "isInherited": false,
        "linkDetails": {
          "AllowsAnonymousAccess": false,
          "ApplicationId": null,
          "BlocksDownload": false,
          "Created": "",
          "CreatedBy": null,
          "Description": null,
          "Embeddable": false,
          "Expiration": "",
          "HasExternalGuestInvitees": false,
          "Invitations": [],
          "IsActive": false,
          "IsAddressBarLink": false,
          "IsCreateOnlyLink": false,
          "IsDefault": true,
          "IsEditLink": false,
          "IsFormsLink": false,
          "IsReviewLink": false,
          "IsUnhealthy": false,
          "LastModified": "",
          "LastModifiedBy": null,
          "LimitUseToApplication": false,
          "LinkKind": 2,
          "PasswordLastModified": "",
          "PasswordLastModifiedBy": null,
          "RedeemedUsers": [],
          "RequiresPassword": false,
          "RestrictedShareMembership": false,
          "Scope": -1,
          "ShareId": "00000000-0000-0000-0000-000000000000",
          "ShareTokenString": null,
          "SharingLinkStatus": 0,
          "TrackLinkUsers": false,
          "Url": null
        },
        "linkMembers": []
      }],
      "principals": [{
        "principal": {
          "email": "",
          "expiration": null,
          "id": 3,
          "isActive": true,
          "isExternal": false,
          "jobTitle": null,
          "loginName": "project-x Owners",
          "name": "project-x Owners",
          "principalType": 8,
          "userId": null,
          "userPrincipalName": null
        },
        "role": 3
      }, {
        "principal": {
          "email": "",
          "expiration": null,
          "id": 4,
          "isActive": true,
          "isExternal": false,
          "jobTitle": null,
          "loginName": "project-x Visitors",
          "name": "project-x Visitors",
          "principalType": 8,
          "userId": null,
          "userPrincipalName": null
        },
        "role": 1
      }, {
        "principal": {
          "email": "",
          "expiration": null,
          "id": 5,
          "isActive": true,
          "isExternal": false,
          "jobTitle": null,
          "loginName": "project-x Members",
          "name": "project-x Members",
          "principalType": 8,
          "userId": null,
          "userPrincipalName": null
        },
        "role": 2
      }],
      "siteAdmins": [{
        "principal": {
          "email": "user1@contoso.onmicrosoft.com",
          "expiration": "",
          "id": 6,
          "isActive": true,
          "isExternal": false,
          "jobTitle": "CEO",
          "loginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com",
          "name": "User 1",
          "principalType": 1,
          "userId": null,
          "userPrincipalName": "user1@contoso.onmicrosoft.com"
        },
        "role": 3
      }]
    }
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_SHARINGINFO_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['fileUrl']);
  });

  it('command correctly handles file get reject request', async () => {
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        fileId: 'f09c4efe-b8c0-4e89-a166-03418661b89b'
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('Retrieves Sharing Information When Site ID is Passed - JSON Output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return {
          "ListItemAllFields": {
            "ParentList": {
              "Title": "Documents"
            },
            "Id": 2,
            "ID": 2
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle') > -1) {
        return JSONOuput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fileId: 'b2307a39-e878-458b-bc90-03bc578531d6',
        output: 'json'
      }
    } as any);
    assert(loggerLogSpy.calledWith(JSONOuput));
  });

  it('Retrieves Sharing Information When document URL is Passed - JSON Output (Debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativePath') > -1) {
        return {
          "ListItemAllFields": {
            "ParentList": {
              "Title": "Documents"
            },
            "Id": 2,
            "ID": 2
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle') > -1) {
        return JSONOuput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fileUrl: '/sites/project-x/documents/SharedFile.docx',
        output: 'json'
      }
    } as any);
    assert(loggerLogToStderrSpy.called);
  });

  it('Retrieves Sharing Information When Site ID is Passed - Text Output (Debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return {
          "ListItemAllFields": {
            "ParentList": {
              "Title": "Documents"
            },
            "Id": 2,
            "ID": 2
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle') > -1) {
        return JSONOuput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fileId: 'b2307a39-e878-458b-bc90-03bc578531d6',
        output: 'text',
        debug: true
      }
    } as any);
    assert(loggerLogToStderrSpy.called);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the fileId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the fileId or fileUrl option not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both fileId and fileUrl options are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'f09c4efe-b8c0-4e89-a166-03418661b89b', fileUrl: '/sites/project-x/documents/SharedFile.docx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
