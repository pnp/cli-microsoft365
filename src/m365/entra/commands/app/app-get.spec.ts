import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { settingsNames } from '../../../../settingsNames.js';
import { telemetry } from '../../../../telemetry.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-get.js';

describe(commands.APP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  //#region Mocked Responses 
  const appResponse = {
    value: [
      {
        "id": "340a4aa3-1af6-43ac-87d8-189819003952"
      }
    ]
  };
  //#endregion

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
      cli.handleMultipleResultsFound,
      entraApp.getAppRegistrationByAppId
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

  it('fails validation when neither appId, objectId, nor name are specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when appId and objectId are both specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
      objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when appId and name are both specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
      name: 'My app'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when objectId and name are both specified', () => {
    const actual = commandOptionsSchema.safeParse({
      objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
      name: 'My app'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when appId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when objectId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      objectId: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when appId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when objectId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when name is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'My app'
    });
    assert.strictEqual(actual.success, true);
  });

  it('handles error when the app specified with the appId not found', async () => {
    const error = `App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      })
    }), new CommandError(`App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`));
  });

  it('handles error when the app with the specified the name not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return { value: [] };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'My app'
      })
    }), new CommandError(`No Microsoft Entra application registration with name My app found`));
  });

  it('handles error when multiple apps with the specified name found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
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
      options: commandOptionsSchema.parse({
        name: 'My app'
      })
    }), new CommandError(`Multiple Microsoft Entra application registrations with name 'My app' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
  });

  it('handles selecting single result when multiple apps with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20App'&$select=id`) {
        return {
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/9b1b1e42-794b-4c71-93ac-5ed92488b67f') {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'My App',
        debug: true
      })
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

  it('handles error when retrieving information about app through name failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'My app'
      })
    }), new CommandError('An error has occurred'));
  });

  it(`should get an Microsoft Entra app registration by its app (client) ID. Doesn't save the app info if not requested`, async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === "https://graph.microsoft.com/v1.0/myorganization/applications/340a4aa3-1af6-43ac-87d8-189819003952?$select=id,appId,displayName") {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "displayName": "My App"
        };
      }

      throw 'Invalid request';
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        properties: 'id,appId,displayName'
      })
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`should get an Microsoft Entra app registration by its name. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20App'&$select=id`) {
        return {
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        };
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
      }

      throw 'Invalid request';
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'My App'
      })
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`should get an Microsoft Entra app registration by its object ID. Doesn't save the app info if not requested`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/340a4aa3-1af6-43ac-87d8-189819003952?$select=id,appId,displayName`) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
      }
      throw 'Invalid request';
    });
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        objectId: '340a4aa3-1af6-43ac-87d8-189819003952',
        properties: 'id,appId,displayName'
      })
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

    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'writeFileSync').callsFake((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      })
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

    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
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
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      })
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

    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
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
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      })
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

    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
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
      options: commandOptionsSchema.parse({
        debug: true,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      })
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
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').throws(new Error('An error has occurred'));
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      })
    });
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save app info in the .m365rc.json file when file has invalid JSON`, async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{');
    const fsWriteFileSyncSpy = sinon.spy(fs, 'writeFileSync');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      })
    });
    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't fail execution when error occurred while saving app info`, async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'writeFileSync').throws(new Error('Error occurred while saving app info'));


    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        save: true
      })
    });
  });
});
