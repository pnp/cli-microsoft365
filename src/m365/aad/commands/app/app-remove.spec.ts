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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-remove');

describe(commands.APP_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  let deleteRequestStub: sinon.SinonStub;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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

    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });

    promptOptions = undefined;

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));

    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        // fake call for getting app
        if (opts.url.indexOf('startswith') === -1) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications(id)",
            "value": [
              {
                "id": "d75be2e1-0204-4f95-857d-51a37cf40be8"
              }
            ]
          };
        }
      }
      throw 'Invalid request';
    });

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
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('prompts before removing the app when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the app when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, {
      options: {
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      }
    });
    assert(deleteRequestStub.notCalled);
  });

  it('deletes app when prompt confirmed (debug)', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        debug: true,
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8'
      }
    });
    assert(deleteRequestStub.called);
  });

  it('deletes app with specified app (client) ID', async () => {
    await command.action(logger, {
      options: {
        appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8',
        confirm: true
      }
    });
    assert(deleteRequestStub.called);
  });

  it('deletes app with specified object ID', async () => {
    await command.action(logger, {
      options: {
        objectId: 'd75be2e1-0204-4f95-857d-51a37cf40be8',
        confirm: true
      }
    });
    assert(deleteRequestStub.called);
  });

  it('deletes app with specified name', async () => {
    await command.action(logger, {
      options: {
        name: 'myapp',
        confirm: true
      }
    });
    assert(deleteRequestStub.called);
  });

  it('fails to get app by id when app does not exists', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        return { value: [] };
      }
      throw "No Azure AD application registration with ID myapp found";
    });

    await assert.rejects(command.action(logger, { options: { debug: true, appId: 'd75be2e1-0204-4f95-857d-51a37cf40be8', confirm: true } } as any), new CommandError("No Azure AD application registration with ID d75be2e1-0204-4f95-857d-51a37cf40be8 found"));
  });

  it('fails to get app by name when app does not exists', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        return { value: [] };
      }
      throw 'No Azure AD application registration with name myapp found';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'myapp', confirm: true } } as any), new CommandError("No Azure AD application registration with name myapp found"));
  });

  it('fails when multiple apps with same name exists', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/myorganization/applications?$filter=`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
          "value": [
            {
              "id": "d75be2e1-0204-4f95-857d-51a37cf40be8"
            },
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952"
            }
          ]
        };
      }

      throw "Multiple Azure AD application registration with name myapp found. Please choose one of the object IDs: d75be2e1-0204-4f95-857d-51a37cf40be8, 340a4aa3-1af6-43ac-87d8-189819003952";
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'myapp',
        confirm: true
      }
    }), new CommandError("Multiple Azure AD application registration with name myapp found. Please choose one of the object IDs: d75be2e1-0204-4f95-857d-51a37cf40be8, 340a4aa3-1af6-43ac-87d8-189819003952"));
  });
});
