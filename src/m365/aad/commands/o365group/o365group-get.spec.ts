import { Group } from '@microsoft/microsoft-graph-types';
import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./o365group-get');

describe(commands.O365GROUP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified Microsoft 365 Group', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve(<Group>{
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

    command.action(logger, { options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('retrieves information about the specified Microsoft 365 Group (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve(<Group>{
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

    command.action(logger, { options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('retrieves information about the specified Microsoft 365 Group including its site URL', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve(<Group>{
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
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('retrieves information about the specified Microsoft 365 Group including its site URL (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve(<Group>{
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
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('retrieves information about the specified Microsoft 365 Group including its site URL (group has no site)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return Promise.resolve(<Group>{
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
        return Promise.resolve(<Group>{
          webUrl: ""
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
