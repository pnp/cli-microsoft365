import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './managementapp-add.js';
import { settingsNames } from '../../../../settingsNames.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { entraApp } from '../../../../utils/entraApp.js';

describe(commands.MANAGEMENTAPP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
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
      request.put,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      entraApp.getAppRegistrationByObjectId,
      entraApp.getAppRegistrationByAppName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MANAGEMENTAPP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles error when the app with the specified the name not found', async () => {
    const error = `App with name 'My app' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My app'
      }
    }), new CommandError(error));
  });

  it('handles error when multiple apps with the specified name found', async () => {
    const error = `Multiple apps with name 'My app' found in Microsoft Entra ID. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My app'
      }
    }), new CommandError(error));
  });

  it('handles error when retrieving information about app through appId failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }), new CommandError(`An error has occurred`));
  });

  it('handles error when retrieving information about app through name failed', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My app'
      }
    }), new CommandError(`An error has occurred`));
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

  it('successfully registers app as managementapp when passing appId', async () => {
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('providers/Microsoft.BusinessAppPlatform/adminApplications/9b1b1e42-794b-4c71-93ac-5ed92488b67f?api-version=2020-06-01') > -1) {
        return {
          "applicationId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].applicationId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
  });

  it('successfully registers app as managementapp when passing name ', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves({
      "id": "340a4aa3-1af6-43ac-87d8-189819003952",
      "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
      "createdDateTime": "2019-10-29T17:46:55Z",
      "displayName": "My Test App",
      "description": null
    });

    sinon.stub(request, 'put').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('providers/Microsoft.BusinessAppPlatform/adminApplications/9b1b1e42-794b-4c71-93ac-5ed92488b67f?api-version=2020-06-01') > -1) {
        return {
          "applicationId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: 'My Test App', debug: true
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].applicationId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
  });

  it('successfully registers app as managementapp when passing objectId ', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByObjectId').resolves({
      "id": "340a4aa3-1af6-43ac-87d8-189819003952",
      "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
      "createdDateTime": "2019-10-29T17:46:55Z",
      "displayName": "My Test App",
      "description": null
    });

    sinon.stub(request, 'put').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('providers/Microsoft.BusinessAppPlatform/adminApplications/9b1b1e42-794b-4c71-93ac-5ed92488b67f?api-version=2020-06-01') > -1) {
        return {
          "applicationId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        objectId: '340a4aa3-1af6-43ac-87d8-189819003952', debug: true
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].applicationId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
  });
});
