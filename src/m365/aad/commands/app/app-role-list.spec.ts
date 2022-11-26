import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-role-list');

describe(commands.APP_ROLE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      telemetry.trackEvent,
      pid.getProcessName
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

  it('lists roles for the specified appId (debug)', async () => {
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

    await command.action(logger, { options: { debug: true, appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8' } });
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
  });

  it('lists roles for the specified appName (debug)', async () => {
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

    await command.action(logger, { options: { debug: true, appName: 'My app' } });
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
  });

  it('lists roles for the specified appId', async () => {
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

    await command.action(logger, { options: { appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230' } });
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
  });

  it(`returns an empty array if the specified app has no roles`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230/appRoles`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230' } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('handles error when the app specified with appObjectId not found', async () => {
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

    await assert.rejects(command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230'
      }
    }), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('handles error when the app specified with the appId not found', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }), new CommandError(`No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`));
  });

  it('handles error when the app specified with appName not found', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app'
      }
    }), new CommandError(`No Azure AD application registration with name My app found`));
  });

  it('handles error when multiple apps with the specified appName found', async () => {
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

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app'
      }
    }), new CommandError(`Multiple Azure AD application registration with name My app found. Please disambiguate (app object IDs): 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g`));
  });

  it('handles error when retrieving information about app through appId failed', async () => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('handles error when retrieving information about app through appName failed', async () => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app'
      }
    } as any), new CommandError('An error has occurred'));
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
});
