import assert from 'assert';
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
import command from './user-recyclebinitem-remove.js';

describe(commands.USER_RECYCLEBINITEM_REMOVE, () => {
  const validUserId = 'd839826a-81bf-4c38-8f80-f150d11ce6c7';

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

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
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    promptOptions = undefined;
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_RECYCLEBINITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the user when prompt confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validUserId}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validUserId, verbose: true } });
    assert(deleteStub.called);
  });

  it('removes the user without prompting the user', async () => {
    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validUserId}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validUserId, force: true, verbose: true } });
    assert(deleteStub.called);
  });

  it('prompts before removing user', async () => {
    await command.action(logger, { options: { id: validUserId } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('aborts removing users when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: validUserId } });
    assert(deleteStub.notCalled);
  });

  it('correctly handles API error', async () => {
    sinon.stub(request, 'delete').rejects({
      error: {
        error: {
          code: 'Request_ResourceNotFound',
          message: `Resource '${validUserId}' does not exist or one of its queried reference-property objects are not present.`,
          innerError: {
            'request-id': '9b0df954-93b5-4de9-8b99-43c204a8aaf8',
            date: '2018-04-24T18:56:48'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { force: true, id: validUserId } } as any),
      new CommandError(`Resource '${validUserId}' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
