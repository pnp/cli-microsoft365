import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./o365group-add');

describe(commands.O365GROUP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.put,
      request.get,
      global.setTimeout
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      fs.readFileSync,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates Microsoft 365 Group using basic info', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group using basic info (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates private Microsoft 365 Group using basic info', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Private'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Private"
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'true' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Private"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group with a png logo', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/png') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' } } as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group with a jpg logo (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/jpeg') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.jpg' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group with a gif logo', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value' &&
        opts.headers &&
        opts.headers['content-type'] === 'image/gif') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.gif' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when creating Microsoft 365 Group with a logo', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when creating Microsoft 365 Group with a logo (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/photo/$value') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group with specific owner', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group with specific owners (debug)', (done) => {
    let groupCreated: boolean = false;
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          groupCreated = true;
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b',
              userPrincipalName: 'user2@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(groupCreated);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group with specific member', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
          deletedDateTime: null,
          classification: null,
          createdDateTime: "2018-02-24T18:38:53Z",
          description: "My awesome group",
          displayName: "My group",
          groupTypes: ["Unified"],
          mail: "my_group@contoso.onmicrosoft.com",
          mailEnabled: true,
          mailNickname: "my_group",
          onPremisesLastSyncDateTime: null,
          onPremisesProvisioningErrors: [],
          onPremisesSecurityIdentifier: null,
          onPremisesSyncEnabled: null,
          preferredDataLocation: null,
          proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
          renewedDateTime: "2018-02-24T18:38:53Z",
          securityEnabled: false,
          visibility: "Public"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft 365 Group with specific members (debug)', (done) => {
    let groupCreated: boolean = false;
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.data) === JSON.stringify({
          description: 'My awesome group',
          displayName: 'My group',
          groupTypes: [
            "Unified"
          ],
          mailEnabled: true,
          mailNickname: 'my_group',
          securityEnabled: false,
          visibility: 'Public'
        })) {
          groupCreated = true;
          return Promise.resolve({
            id: "f3db5c2b-068f-480d-985b-ec78b9fa0e76",
            deletedDateTime: null,
            classification: null,
            createdDateTime: "2018-02-24T18:38:53Z",
            description: "My awesome group",
            displayName: "My group",
            groupTypes: ["Unified"],
            mail: "my_group@contoso.onmicrosoft.com",
            mailEnabled: true,
            mailNickname: "my_group",
            onPremisesLastSyncDateTime: null,
            onPremisesProvisioningErrors: [],
            onPremisesSecurityIdentifier: null,
            onPremisesSyncEnabled: null,
            preferredDataLocation: null,
            proxyAddresses: ["SMTP:my_group@contoso.onmicrosoft.com"],
            renewedDateTime: "2018-02-24T18:38:53Z",
            securityEnabled: false,
            visibility: "Public"
          });
        }
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.data['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b',
              userPrincipalName: 'user2@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(groupCreated);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when an invalid user is specified as owner', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when an invalid user is specified as owner (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when an invalid user is specified as member', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when an invalid user is specified as member (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user1@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } }, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Cannot proceed with group creation. The following users provided are invalid : user2@contoso.onmicrosoft.com");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6320', scope: 'user_impersonation' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation when the displayName, description and mailNickname are specified', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if one of the owners is invalid', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the owner is valid', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple owners, comma-separated', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple owners, comma-separated with an additional space', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if one of the members is invalid', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the member is valid', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple members, comma-separated', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation with multiple members, comma-separated with an additional space', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if isPrivate is invalid boolean', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if isPrivate is true', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if isPrivate is false', () => {
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'false' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if logoPath points to a non-existent file', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'invalid' } });
    sinonUtil.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if logoPath points to a folder', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'folder' } });
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if logoPath points to an existing file', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = command.validate({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'folder' } });
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying mailNickname', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--mailNickname') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying owners', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--owners') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying members', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--members') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying group type', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--isPrivate') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying logo file path', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--logoPath') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});