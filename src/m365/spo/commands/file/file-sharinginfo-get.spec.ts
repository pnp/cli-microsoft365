import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./file-sharinginfo-get');

describe(commands.FILE_SHARINGINFO_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_SHARINGINFO_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('command correctly handles file get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 'f09c4efe-b8c0-4e89-a166-03418661b89b'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('Retrieves Sharing Information When Site ID is Passed - JSON Output', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "ParentList": {
              "Title": "Documents"
            },
            "Id": 2,
            "ID": 2
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle') > -1) {
        return Promise.resolve(JSONOuput);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        output: 'json'
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(JSONOuput));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Retrieves Sharing Information When document URL is Passed - JSON Output (Debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativePath') > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "ParentList": {
              "Title": "Documents"
            },
            "Id": 2,
            "ID": 2
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle') > -1) {
        return Promise.resolve(JSONOuput);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        url: '/sites/project-x/documents/SharedFile.docx',
        output: 'json'
      }
    } as any, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Retrieves Sharing Information When Site ID is Passed - Text Output (Debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "ParentList": {
              "Title": "Documents"
            },
            "Id": 2,
            "ID": 2
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle') > -1) {
        return Promise.resolve(JSONOuput);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        output: 'text',
        debug: true
      }
    } as any, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert(actual);
  });

  it('fails validation if the id or url option not specified', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and url options are specified', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', url: '/sites/project-x/documents/SharedFile.docx' } });
    assert.notStrictEqual(actual, true);
  });
});