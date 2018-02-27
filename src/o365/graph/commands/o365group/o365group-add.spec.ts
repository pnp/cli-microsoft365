import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./o365group-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';
import * as fs from 'fs';

describe(commands.O365GROUP_ADD, () => {
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
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
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
      request.post,
      request.put,
      request.get,
      global.setTimeout
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth,
      fs.readFileSync
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.O365GROUP_ADD), true);
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
        assert.equal(telemetry.name, commands.O365GROUP_ADD);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to the Microsoft Graph', (done) => {
    auth.service = new Service('https://graph.microsoft.com');
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

  it('creates Office 365 Group using basic info', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates Office 365 Group using basic info (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates private Office 365 Group using basic info', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates Office 365 Group with a png logo', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
        opts.headers['content-type'] === 'image/png') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates Office 365 Group with a jpg logo (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
        opts.headers['content-type'] === 'image/jpeg') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.jpg' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates Office 365 Group with a gif logo', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
        opts.headers['content-type'] === 'image/gif') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.gif' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('handles failure when creating Office 365 Group with a logo', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when creating Office 365 Group with a logo (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
    sinon.stub(global, 'setTimeout').callsFake((fn, to) => {
      fn();
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'logo.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Office 365 Group with specific owner', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates Office 365 Group with specific owners (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/owners/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        })
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates Office 365 Group with specific member', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('creates Office 365 Group with specific members (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        if (JSON.stringify(opts.body) === JSON.stringify({
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
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8a') {
        return Promise.resolve();
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/f3db5c2b-068f-480d-985b-ec78b9fa0e76/members/$ref' &&
        opts.body['@odata.id'] === 'https://graph.microsoft.com/v1.0/users/949b16c1-a032-453e-a8ae-89a52bfc1d8b') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user1@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
            }
          ]
        })
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2@contoso.onmicrosoft.com'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8b'
            }
          ]
        })
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user@contoso.onmicrosoft.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
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

    auth.service = new Service('https://graph.windows.net');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6320', scope: 'user_impersonation' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the displayName is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { description: 'My awesome group', mailNickname: 'my_group' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the description is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', mailNickname: 'my_group' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the mailNickname is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the displayName, description and mailNickname are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group' } });
    assert.equal(actual, true);
  });

  it('fails validation if one of the owners is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the owner is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user@contoso.onmicrosoft.com' } });
    assert.equal(actual, true);
  });

  it('passes validation with multiple owners, comma-separated', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert.equal(actual, true);
  });

  it('passes validation with multiple owners, comma-separated with an additional space', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', owners: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } });
    assert.equal(actual, true);
  });

  it('fails validation if one of the members is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the member is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user@contoso.onmicrosoft.com' } });
    assert.equal(actual, true);
  });

  it('passes validation with multiple members, comma-separated', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com' } });
    assert.equal(actual, true);
  });

  it('passes validation with multiple members, comma-separated with an additional space', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', members: 'user1@contoso.onmicrosoft.com, user2@contoso.onmicrosoft.com' } });
    assert.equal(actual, true);
  });

  it('fails validation if isPrivate is invalid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if isPrivate is true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation if isPrivate is false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', isPrivate: 'false' } });
    assert.equal(actual, true);
  });

  it('fails validation if logoPath points to a non-existent file', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'invalid' } });
    Utils.restore(fs.existsSync);
    assert.notEqual(actual, true);
  });

  it('fails validation if logoPath points to a folder', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'folder' } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notEqual(actual, true);
  });

  it('passes validation if logoPath points to an existing file', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { displayName: 'My group', description: 'My awesome group', mailNickname: 'my_group', logoPath: 'folder' } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
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

  it('supports specifying displayName', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying mailNickname', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--mailNickname') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying owners', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--owners') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying members', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--members') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying group type', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--isPrivate') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying logo file path', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--logoPath') > -1) {
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
    assert(find.calledWith(commands.O365GROUP_ADD));
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