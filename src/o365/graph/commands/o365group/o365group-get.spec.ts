import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./o365group-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.O365GROUP_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
    auth.service = new Service('https://graph.microsoft.com');
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.O365GROUP_GET), true);
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
        assert.equal(telemetry.name, commands.O365GROUP_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to Microsoft Graph', (done) => {
    auth.service = new Service();
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to the Microsoft Graph first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });



  it('retrieves information about the specified Office 365 Group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified Office 365 Group (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified Office 365 Group including its site URL', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "siteUrl": "https://contoso.sharepoint.com/sites/finance",
          "visibility": "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified Office 365 Group including its site URL (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "siteUrl": "https://contoso.sharepoint.com/sites/finance",
          "visibility": "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified Office 365 Group including its site URL (group has no site)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: ""
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public",
          "siteUrl": ""
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no group found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c843`) {
        return Promise.reject({
          error: {
            "error": {
              "code": "Request_ResourceNotFound",
              "message": "Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.",
              "innerError": {
                "request-id": "7e192558-7438-46db-a4c9-5dca83d2ec96",
                "date": "2018-02-21T20:38:50"
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c843' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(`Resource '1caf7dcd-7e83-4c3a-94f7-932a1299c843' does not exist or one of its queried reference-property objects are not present.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the id is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '123' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert.equal(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
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
    assert(find.calledWith(commands.O365GROUP_GET));
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
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});