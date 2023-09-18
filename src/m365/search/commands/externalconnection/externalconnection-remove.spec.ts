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
import command from './externalconnection-remove.js';

describe(commands.EXTERNALCONNECTION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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

    promptOptions = undefined;
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EXTERNALCONNECTION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the specified external connection by id when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        id: "contosohr"
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the specified external connection by name when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        name: "Contoso HR"
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified external connection when confirm option not passed and prompt not confirmed (debug)', async () => {
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

    sinonUtil.restore(Cli.prompt);
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

      throw 'The specified connection does not exist in Microsoft Search';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: "Fabrikam HR",
        force: true
      }
    } as any), new CommandError("The specified connection does not exist in Microsoft Search"));
  });

  it('fails when multiple external connections with same name exists', async () => {
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
