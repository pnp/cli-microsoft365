import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-label-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_LABEL_SET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.LIST_LABEL_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should handle error when trying to set label', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url == `https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) {
        return Promise.reject({
          error: {
            'odata.error': {
              code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
              message: {
                value: 'Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4'
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder`) {
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
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle error if list does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder`) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        label: 'abc'
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

  it('should set label for list (debug)', (done) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder`) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        label: 'abc'
      }
    }, () => {
      try {
        const lastCall = postStub.lastCall.args[0];
        assert.strictEqual(lastCall.body.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
        assert.strictEqual(lastCall.body.complianceTagValue, 'abc');
        assert.strictEqual(lastCall.body.blockDelete, false);
        assert.strictEqual(lastCall.body.blockEdit, false);
        assert.strictEqual(lastCall.body.syncToItems, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should set label for list using listId (debug)', (done) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'4d535433-2a7b-40b0-9dad-8f0f8f3b3841')/?$expand=RootFolder&$select=RootFolder`) {
        return Promise.resolve({ "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } }
        );
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: '4d535433-2a7b-40b0-9dad-8f0f8f3b3841',
        label: 'abc'
      }
    }, () => {
      try {
        const lastCall = postStub.lastCall.args[0];
        assert.strictEqual(lastCall.body.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
        assert.strictEqual(lastCall.body.complianceTagValue, 'abc');
        assert.strictEqual(lastCall.body.blockDelete, false);
        assert.strictEqual(lastCall.body.blockEdit, false);
        assert.strictEqual(lastCall.body.syncToItems, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should set label for list using listUrl option (verbose)', (done) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listUrl: 'MyLibrary',
        label: 'abc'
      }
    }, () => {
      try {
        const lastCall = postStub.lastCall.args[0];
        assert.strictEqual(lastCall.body.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
        assert.strictEqual(lastCall.body.complianceTagValue, 'abc');
        assert.strictEqual(lastCall.body.blockDelete, false);
        assert.strictEqual(lastCall.body.blockEdit, false);
        assert.strictEqual(lastCall.body.syncToItems, false);
        assert.notStrictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should set label for list using blockDelete,blockEdit,syncToItems options', (done) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listUrl: 'MyLibrary',
        label: 'abc',
        blockDelete: true,
        blockEdit: true,
        syncToItems: true
      }
    }, () => {
      try {
        const lastCall = postStub.lastCall.args[0];
        assert.strictEqual(lastCall.body.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
        assert.strictEqual(lastCall.body.complianceTagValue, 'abc');
        assert.strictEqual(lastCall.body.blockDelete, true);
        assert.strictEqual(lastCall.body.blockEdit, true);
        assert.strictEqual(lastCall.body.syncToItems, true);
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
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', label: 'abc', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', label: 'abc', listId: 'XXXXX' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', label: 'abc', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert(actual);
  });

  it('fails validation if listId, listUrl and listTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', label: 'abc' } });
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