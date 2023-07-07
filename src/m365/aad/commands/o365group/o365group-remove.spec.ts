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
const command: Command = require('./o365group-remove');

describe(commands.O365GROUP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
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
    loggerLogSpy = sinon.spy(logger, 'log');
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      global.setTimeout,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.O365GROUP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified group without prompting for confirmation when confirm option specified', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', confirm: false } });
    assert(loggerLogSpy.notCalled);
  });

  it('removes the specified group without prompting for confirmation when confirm option specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848', confirm: false } });
    assert(loggerLogSpy.notCalled);
  });

  it('prompts before removing the specified group when confirm option not passed', async () => {
    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the specified group when confirm option not passed (debug)', async () => {
    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the group when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(postSpy.notCalled);
  });

  it('aborts removing the group when prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(postSpy.notCalled);
  });

  it('removes the group when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'delete').resolves();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(postStub.called);
  });

  it('removes the group when prompt confirmed (debug)', async () => {
    const postStub = sinon.stub(request, 'delete').resolves();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(postStub.called);
  });

  it('removes the group permanently when prompt confirmed', async () => {
    let groupPermDeleteCallIssued = false;
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/28beab62-7540-4db1-a23f-29a6018a3848`) {
        groupPermDeleteCallIssued = true;
        return;
      }

      throw 'Invalid request';
    });
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848', skipRecycleBin: true } });
    assert(groupPermDeleteCallIssued);
  });

  it('removes the group permanently when with confirm option', async () => {
    let groupPermDeleteCallIssued = false;
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848`) {
        return;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/28beab62-7540-4db1-a23f-29a6018a3848`) {
        groupPermDeleteCallIssued = true;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848', skipRecycleBin: true, confirm: true } });
    assert(groupPermDeleteCallIssued);
  });

  it('correctly handles error when group is not found', async () => {
    sinon.stub(request, 'delete').rejects({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });

    await assert.rejects(command.action(logger, { options: { confirm: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } } as any),
      new CommandError('File Not Found.'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirmation flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--confirm') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
