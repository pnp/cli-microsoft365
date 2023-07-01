import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './groupsetting-remove.js';

describe(commands.GROUPSETTING_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(fs, 'readFileSync').returns('abc');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
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
    assert.strictEqual(command.name, commands.GROUPSETTING_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified group setting without prompting for confirmation when confirm option specified', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groupSettings/28beab62-7540-4db1-a23f-29a6018a3848') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848', force: true } });
    assert(deleteRequestStub.called);
  });

  it('removes the specified group setting without prompting for confirmation when confirm option specified (debug)', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groupSettings/28beab62-7540-4db1-a23f-29a6018a3848') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848', force: true } });
    assert(deleteRequestStub.called);
  });

  it('prompts before removing the specified group setting when confirm option not passed', async () => {
    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('prompts before removing the specified group setting when confirm option not passed (debug)', async () => {
    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the group setting when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(postSpy.notCalled);
  });

  it('aborts removing the group setting when prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(postSpy.notCalled);
  });

  it('removes the group setting when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'delete').callsFake(() => Promise.resolve());

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(postStub.called);
  });

  it('removes the group setting when prompt confirmed (debug)', async () => {
    const deleteStub = sinon.stub(request, 'delete').resolves();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } });
    assert(deleteStub.called);
  });

  it('correctly handles error when group setting is not found', async () => {
    sinon.stub(request, 'delete').rejects({
      error: { 'odata.error': { message: { value: 'File Not Found.' } } }
    });

    await assert.rejects(command.action(logger, { options: { force: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } } as any),
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
      if (o.option.indexOf('--force') > -1) {
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
