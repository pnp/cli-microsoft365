import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { settingsNames } from '../../../../settingsNames.js';
import { telemetry } from '../../../../telemetry.js';
import { misc } from '../../../../utils/misc.js';
import { MockRequests } from '../../../../utils/MockRequest.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';
import command from './app-get.js';

export const mocks = {
  getById: {
    request: {
      url: `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`
    },
    response: {
      body: {
        value: [
          {
            "id": "340a4aa3-1af6-43ac-87d8-189819003952",
            "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            "createdDateTime": "2019-10-29T17:46:55Z",
            "displayName": "My App",
            "description": null
          }
        ]
      }
    }
  },
  getByName: {
    request: {
      url: `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20App'&$select=id`
    },
    response: {
      body: {
        value: [
          {
            "id": "340a4aa3-1af6-43ac-87d8-189819003952",
            "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
            "createdDateTime": "2019-10-29T17:46:55Z",
            "displayName": "My App",
            "description": null
          }
        ]
      }
    }
  },
  getAppByAppId: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/myorganization/applications/9b1b1e42-794b-4c71-93ac-5ed92488b67f'
    },
    response: {
      body: {
        "id": "340a4aa3-1af6-43ac-87d8-189819003952",
        "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
        "createdDateTime": "2019-10-29T17:46:55Z",
        "displayName": "My App",
        "description": null
      }
    }
  },
  getAppByObjectId: {
    request: {
      url: `https://graph.microsoft.com/v1.0/myorganization/applications/340a4aa3-1af6-43ac-87d8-189819003952`
    },
    response: {
      body: {
        "id": "340a4aa3-1af6-43ac-87d8-189819003952",
        "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
        "createdDateTime": "2019-10-29T17:46:55Z",
        "displayName": "My App",
        "description": null
      }
    }
  }
} satisfies MockRequests;

describe(commands.APP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
      fs.writeFileSync,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.APP_GET, commands.APPREGISTRATION_GET]);
  });

  it('handles error when the app specified with the appId not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === mocks.getById.request.url) {
        return { value: [] };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }), new CommandError(`No Microsoft Entra application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`));
  });

  it('handles error when the app with the specified the name not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === mocks.getByName.request.url) {
        return { value: [] };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My App'
      }
    }), new CommandError(`No Microsoft Entra application registration with name My App found`));
  });

  it('handles error when multiple apps with the specified name found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === mocks.getByName.request.url) {
        return {
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My App'
      }
    }), new CommandError(`Multiple Microsoft Entra application registrations with name 'My App' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
  });

  it('handles selecting single result when multiple apps with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === mocks.getByName.request.url) {
        return {
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      if (opts.url === mocks.getAppByAppId.request.url) {
        return misc.deepClone(mocks.getAppByAppId.response.body);
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });

    await command.action(logger, {
      options: {
        name: 'My App'
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.deepEqual(call.args[0], {
      "id": "340a4aa3-1af6-43ac-87d8-189819003952",
      "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
      "createdDateTime": "2019-10-29T17:46:55Z",
      "displayName": "My App",
      "description": null
    });
  });

  it('handles error when retrieving information about app through appId failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('handles error when retrieving information about app through name failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My app'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if appId and objectId specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and name specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and name specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, objectId, nor name specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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

  it(`should get an Microsoft Entra app registration by its app (client) ID. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      throw 'Invalid request';
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

  it(`should get an Microsoft Entra app registration by its name. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getByName.request.url) {
        return misc.deepClone(mocks.getByName.response.body);
      }

      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      throw 'Invalid request';
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

  it(`should get an Microsoft Entra app registration by its object ID. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }
      throw 'Invalid request';
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

  it(`should get an Microsoft Entra app registration by its app (client) ID. Creates the file it doesn't exist`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(false);
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

  it(`should get an Microsoft Entra app registration by its app (client) ID. Writes to the existing empty file`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('');
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

  it(`should get an Microsoft Entra app registration by its app (client) ID. Adds to the existing file contents`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
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

  it(`should get an Microsoft Entra app registration by its app (client) ID. Adds to the existing file contents (Debug)`, async () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').throws(new Error('An error has occurred'));
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{');
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === mocks.getAppByObjectId.request.url) {
        return misc.deepClone(mocks.getAppByObjectId.response.body);
      }

      if (opts.url === mocks.getById.request.url) {
        return misc.deepClone(mocks.getById.response.body);
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'writeFileSync').throws(new Error('Error occurred while saving app info'));


    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      }
    });
  });
});
