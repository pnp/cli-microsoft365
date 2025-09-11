import assert from 'assert';
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
import command from './app-remove.js';

describe(commands.APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let promptIssued: boolean = false;
  let deleteRequestStub: sinon.SinonStub;

  //#region Mocked Responses 
  const appResponse = {
    value: [
      {
        "id": "d75be2e1-0204-4f95-857d-51a37cf40be8"
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

    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;

    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);
    sinon.stub(entraApp, 'getAppRegistrationByAppName').resolves(appResponse.value[0]);

    deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/d75be2e1-0204-4f95-857d-51a37cf40be8') {
        return;
      }
      throw 'Invalid request';
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      entraApp.getAppRegistrationByAppId,
      entraApp.getAppRegistrationByAppName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if appId and name specified', () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if objectId and name specified', () => {
    const actual = commandOptionsSchema.safeParse({ objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither appId, objectId, nor name specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the objectId is not a valid guid', () => {
    const actual = commandOptionsSchema.safeParse({ objectId: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the appId is not a valid guid', () => {
    const actual = commandOptionsSchema.safeParse({ appId: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if required options specified (appId)', () => {
    const actual = commandOptionsSchema.safeParse({ appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required options specified (objectId)', () => {
    const actual = commandOptionsSchema.safeParse({ objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required options specified (name)', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'My app' });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing the app when force option not passed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      })
    });

    assert(promptIssued);
  });

  it('aborts removing the app when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      })
    });
    assert(deleteRequestStub.notCalled);
  });

  it('deletes app when prompt confirmed (debug)', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      })
    });
    assert(deleteRequestStub.called);
  });

  it('deletes app with specified app (client) ID', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8',
        force: true
      })
    });
    assert(deleteRequestStub.called);
  });

  it('deletes app with specified object ID', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        objectId: 'd75be2e1-0204-4f95-857d-51a37cf40be8',
        force: true
      })
    });
    assert(deleteRequestStub.called);
  });

  it('deletes app with specified name', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'myapp',
        force: true
      })
    });
    assert(deleteRequestStub.called);
  });

  it('fails to get app by id when app does not exists', async () => {
    sinonUtil.restore(entraApp.getAppRegistrationByAppId);
    const error = `App with appId 'd75be2e1-0204-4f95-857d-51a37cf40be8' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8', force: true }) }), new CommandError(error));
  });

  it('fails to get app by name when app does not exists', async () => {
    sinonUtil.restore(entraApp.getAppRegistrationByAppName);
    const error = `App with name 'myapp' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'myapp', force: true }) }), new CommandError(error));
  });

  it('fails when multiple apps with same name exists', async () => {
    sinonUtil.restore(entraApp.getAppRegistrationByAppName);
    const error = `Multiple apps with name 'myapp' found in Microsoft Entra ID. Found: d75be2e1-0204-4f95-857d-51a37cf40be8, 340a4aa3-1af6-43ac-87d8-189819003952.`;
    sinon.stub(entraApp, 'getAppRegistrationByAppName').rejects(new Error(error));
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        name: 'myapp',
        force: true
      })
    }), new CommandError(error));
  });

  it('handles selecting single result when multiple apps with the specified name found and cli is set to prompt', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'myapp'&$select=id`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
          "value": [
            { "id": "d75be2e1-0204-4f95-857d-51a37cf40be8" },
            { "id": "340a4aa3-1af6-43ac-87d8-189819003952" }
          ]
        };
      }

      throw "Multiple Microsoft Entra application registration with name 'myapp' found.";
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: 'd75be2e1-0204-4f95-857d-51a37cf40be8' });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'myapp',
        force: true
      })
    });
    assert(deleteRequestStub.called);
  });
});
