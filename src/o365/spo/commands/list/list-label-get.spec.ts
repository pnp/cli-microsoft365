import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./list-label-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.LIST_LABEL_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LIST_LABEL_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.LIST_LABEL_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team1', title: 'Documents' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the label from the given list if title option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
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
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the label from the given list if title option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/`) > -1) {
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
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
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the label from the given list if list id option is passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
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
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets the label from the given list if list id option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
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
        assert.equal(JSON.stringify(actual), JSON.stringify(expected));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles the case when no label has been set on the specified list', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'fb4b0cf8-c006-4802-a1ea-57e0e4852188')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists/GetByTitle('MyLibrary')`) > -1) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when trying to get label from a list that doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`) > -1) {
        return Promise.resolve([]);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')`) > -1) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('404 - "404 FILE NOT FOUND"')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'XXXXX' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', listTitle: 'Documents' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both listId and listTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.LIST_LABEL_GET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});