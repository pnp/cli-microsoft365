import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./externalconnection-remove');

describe(commands.EXTERNALCONNECTION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
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

    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.EXTERNALCONNECTION_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['id', 'name'] }]);
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

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        externalConnectionRemoveCallIssued = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));


    await command.action(logger, { options: { debug: true, id: "contosohr" } });
    assert(externalConnectionRemoveCallIssued);
  });

  it('removes the specified external connection without prompting when confirm specified', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, id: "contosohr", confirm: true } });
  });

  it('removes external connection with specified ID', async () => {
    sinon.stub(request, 'delete').callsFake((opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return Promise.resolve();
      }
      return Promise.reject();
    });

    await command.action(logger, { options: { debug: false, id: "contosohr", confirm: true } });
  });

  it('removes external connection with specified name', async () => {
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq `) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "contosohr",
              "name": "Contoso HR",
              "description": "Connection to index Contoso HR system"
            }
          ]
        });
      }
      return Promise.reject();
    });

    sinon.stub(request, 'delete').callsFake((opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return Promise.resolve();
      }
      return Promise.reject();
    });

    await command.action(logger, { options: { debug: false, name: "Contoso HR", confirm: true } });
  });

  it('fails to get external connection by name when it does not exists', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('The specified connection does not exist in Microsoft Search');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        name: "Fabrikam HR",
        confirm: true
      }
    } as any), new CommandError("The specified connection does not exist in Microsoft Search"));
  });

  it('fails when multiple external connections with same name exists', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "fabrikamhr"
            },
            {
              "id": "contosohr"
            }
          ]
        });
      }

      return Promise.reject("Invalid request");
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        name: "My HR",
        confirm: true
      }
    } as any), new CommandError("Multiple external connections with name My HR found. Please disambiguate (IDs): fabrikamhr, contosohr"));
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
