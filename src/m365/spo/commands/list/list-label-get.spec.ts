import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-label-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_LABEL_GET, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.LIST_LABEL_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets the label from the given list if title option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary'
      }
    }, () => {
      try {
        const expected = {
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the label from the given list if title option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/`) > -1) {
        return Promise.resolve({
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary'
      }
    }, () => {
      try {
        const expected = {
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the label from the given list if list id option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'
      }
    }, () => {
      try {
        const expected = {
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null

        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the label from the given list if list id option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188'
      }
    }, () => {
      try {
        const expected = {
          "AcceptMessagesOnlyFromSendersOrMembers": false,
          "AccessType": null,
          "AllowAccessFromUnmanagedDevice": null,
          "AutoDelete": false,
          "BlockDelete": false,
          "BlockEdit": false,
          "ContainsSiteLabel": false,
          "DisplayName": "",
          "EncryptionRMSTemplateId": null,
          "HasRetentionAction": false,
          "IsEventTag": false,
          "Notes": null,
          "RequireSenderAuthenticationEnabled": false,
          "ReviewerEmail": null,
          "SharingCapabilities": null,
          "SuperLock": false,
          "TagDuration": 0,
          "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
          "TagName": "Sensitive",
          "TagRetentionBasedOn": null
        };
        const actual = log[log.length - 1];
        assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles the case when no label has been set on the specified list', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'fb4b0cf8-c006-4802-a1ea-57e0e4852188',
      }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to get label for the list', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.reject({
          error: {
            'odata.error': {
              code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
              message: {
                value: 'An error has occurred'
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to get label from a list that doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve([]);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')`) > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('404 - "404 FILE NOT FOUND"')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'XXXXX' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', listTitle: 'Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both listId and listTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});