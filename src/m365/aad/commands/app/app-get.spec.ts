import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-get');

describe(commands.APP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      request.get,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('handles error when the app with the specified the name not found', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My app'
      }
    }), new CommandError(`No Azure AD application registration with name My app found`));
  });

  it('handles error when multiple apps with the specified name found', async () => {
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
        name: 'My app'
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

  it('handles error when retrieving information about app through name failed', async () => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My app'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if appId and objectId specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and name specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and name specified', async () => {
    const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, objectId, nor name specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid guid', async () => {
    const actual = await command.validate({ options: { objectId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid guid', async () => {
    const actual = await command.validate({ options: { appId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (objectId)', async () => {
    const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { name: 'My app' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`should get an Azure AD app registration by its app (client) ID. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`should get an Azure AD app registration by its name. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20App'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: {
        name: 'My App'
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`should get an Azure AD app registration by its object ID. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/340a4aa3-1af6-43ac-87d8-189819003952`) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }
      return Promise.reject('Invalid request');
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: {
        objectId: '340a4aa3-1af6-43ac-87d8-189819003952'
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`should get an Azure AD app registration by its app (client) ID. Creates the file it doesn't exist`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      apps: [{
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'My App'
      }]
    }, null, 2));
  });

  it(`should get an Azure AD app registration by its app (client) ID. Writes to the existing empty file`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '');
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      apps: [{
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'My App'
      }]
    }, null, 2));
  });

  it(`should get an Azure AD app registration by its app (client) ID. Adds to the existing file contents`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "74ad36da-3704-4e67-ba08-8c8e833f3c52",
          "name": "M365 app"
        }
      ]
    }));
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      apps: [
        {
          "appId": "74ad36da-3704-4e67-ba08-8c8e833f3c52",
          "name": "M365 app"
        },
        {
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          name: 'My App'
        }]
    }, null, 2));
  });

  it(`should get an Azure AD app registration by its app (client) ID. Adds to the existing file contents (Debug)`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => JSON.stringify({
      "apps": [
        {
          "appId": "74ad36da-3704-4e67-ba08-8c8e833f3c52",
          "name": "M365 app"
        }
      ]
    }));
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      apps: [
        {
          "appId": "74ad36da-3704-4e67-ba08-8c8e833f3c52",
          "name": "M365 app"
        },
        {
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          name: 'My App'
        }]
    }, null, 2));
  });

  it(`doesn't save app info in the .m365rc.json file when there was error reading file contents`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => { throw new Error('An error has occurred'); });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save app info in the .m365rc.json file when file has invalid JSON`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '{');
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't fail execution when error occurred while saving app info`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { throw new Error('Error occurred while saving app info'); });


    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
  });
});
