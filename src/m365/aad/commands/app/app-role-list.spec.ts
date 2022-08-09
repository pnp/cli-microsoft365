import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./app-role-list');

describe(commands.APP_ROLE_LIST, () => {
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
    (command as any).items = [];
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
    assert.strictEqual(command.name.startsWith(commands.APP_ROLE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['displayName', 'description', 'id']);
  });

  it('lists roles for the specified appId (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq 'bc724b77-da87-43a9-b385-6ebaaf969db8'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return Promise.resolve({
          value: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Readers",
              "displayName": "Readers",
              "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
              "isEnabled": true,
              "origin": "Application",
              "value": "readers"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Writers",
              "displayName": "Writers",
              "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
              "isEnabled": true,
              "origin": "Application",
              "value": "writers"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Readers",
            "displayName": "Readers",
            "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
            "isEnabled": true,
            "origin": "Application",
            "value": "readers"
          },
          {
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Writers",
            "displayName": "Writers",
            "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
            "isEnabled": true,
            "origin": "Application",
            "value": "writers"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists roles for the specified appName (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return Promise.resolve({
          value: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Readers",
              "displayName": "Readers",
              "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
              "isEnabled": true,
              "origin": "Application",
              "value": "readers"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Writers",
              "displayName": "Writers",
              "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
              "isEnabled": true,
              "origin": "Application",
              "value": "writers"
            }
          ]
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    command.action(logger, { options: { debug: true, appName: 'My app' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Readers",
            "displayName": "Readers",
            "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
            "isEnabled": true,
            "origin": "Application",
            "value": "readers"
          },
          {
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Writers",
            "displayName": "Writers",
            "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
            "isEnabled": true,
            "origin": "Application",
            "value": "writers"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists roles for the specified appId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return Promise.resolve({
          value: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Readers",
              "displayName": "Readers",
              "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
              "isEnabled": true,
              "origin": "Application",
              "value": "readers"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Writers",
              "displayName": "Writers",
              "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
              "isEnabled": true,
              "origin": "Application",
              "value": "writers"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Readers",
            "displayName": "Readers",
            "id": "ca12d0da-cd83-4dc9-8e4c-b6a529bebbb4",
            "isEnabled": true,
            "origin": "Application",
            "value": "readers"
          },
          {
            "allowedMemberTypes": [
              "User"
            ],
            "description": "Writers",
            "displayName": "Writers",
            "id": "85c03d41-b438-48ea-bccd-8389c0e327bc",
            "isEnabled": true,
            "origin": "Application",
            "value": "writers"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`returns an empty array if the specified app has no roles`, (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with appObjectId not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles') {
        return Promise.reject({
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2021-04-20T17:22:30",
              "request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d",
              "client-request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d"
            }
          }
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with the appId not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the app specified with appName not found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appName: 'My app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `No Azure AD application registration with name My app found`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when multiple apps with the specified appName found', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        appName: 'My app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `Multiple Azure AD application registration with name My app found. Please disambiguate (app object IDs): 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through appId failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));

    command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving information about app through appName failed', (done) => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));

    command.action(logger, {
      options: {
        debug: false,
        appName: 'My app'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, `An error has occurred`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if appId and appObjectId specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appObjectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appName specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appName specified', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, appObjectId nor appName specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if appId specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if appObjectId specified', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if appName specified', async () => {
    const actual = await command.validate({ options: { appName: 'My app' } }, commandInfo);
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
});