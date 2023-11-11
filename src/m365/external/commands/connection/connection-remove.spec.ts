import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './connection-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.CONNECTION_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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

    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.getSettingWithDefaultValue,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONNECTION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('prompts before removing the specified external connection by id when force option not passed', async () => {
    await command.action(logger, {
      options: {
        id: "contosohr"
      }
    });

    assert(promptIssued);
  });

  it('prompts before removing the specified external connection by name when force option not passed', async () => {
    await command.action(logger, {
      options: {
        name: "Contoso HR"
      }
    });

    assert(promptIssued);
  });

  it('aborts removing the specified external connection when force option not passed and prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    await command.action(logger, { options: { debug: true, id: "contosohr" } });
    assert(postSpy.notCalled);
  });

  it('removes the specified external connection when prompt confirmed (debug)', async () => {
    let externalConnectionRemoveCallIssued = false;

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        externalConnectionRemoveCallIssued = true;
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);


    await command.action(logger, { options: { debug: true, id: "contosohr" } });
    assert(externalConnectionRemoveCallIssued);
  });

  it('removes the specified external connection without prompting when confirm specified', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: "contosohr", force: true } });
  });

  it('removes external connection with specified ID', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return;
      }
      throw '';
    });

    await command.action(logger, { options: { id: "contosohr", force: true } });
  });

  it('removes external connection with specified name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq `) > -1) {
        return {
          value: [
            {
              "id": "contosohr",
              "name": "Contoso HR",
              "description": "Connection to index Contoso HR system"
            }
          ]
        };
      }
      throw '';
    });

    sinon.stub(request, 'delete').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return;
      }
      throw '';
    });

    await command.action(logger, { options: { name: "Contoso HR", force: true } });
  });

  it('fails to get external connection by name when it does not exists', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
        return { value: [] };
      }

      throw 'The specified connection does not exist';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: "Fabrikam HR",
        force: true
      }
    } as any), new CommandError("The specified connection does not exist"));
  });

  it('fails when multiple external connections with same name exists', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
        return {
          value: [
            {
              "id": "fabrikamhr"
            },
            {
              "id": "contosohr"
            }
          ]
        };
      }

      throw "Invalid request";
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: "My HR",
        force: true
      }
    } as any), new CommandError("Multiple external connections with name My HR found. Found: fabrikamhr, contosohr."));
  });

  it('handles selecting single result when external connections with the specified name found and cli is set to prompt', async () => {
    let removeRequestIssued = false;

    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections?$filter=name eq 'My%20HR'&$select=id`) {
        return {
          value: [
            {
              "id": "fabrikamhr"
            },
            {
              "id": "contosohr"
            }
          ]
        };
      }

      throw "Invalid request";
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves({
      "id": "contosohr"
    });

    sinon.stub(request, 'delete').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        removeRequestIssued = true;
        return;
      }
      throw '';
    });

    await command.action(logger, { options: { name: "My HR", force: true } });
    assert(removeRequestIssued);
  });
});
