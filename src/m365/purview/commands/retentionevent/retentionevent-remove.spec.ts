import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './retentionevent-remove.js';

describe(commands.RETENTIONEVENT_REMOVE, () => {
  const validId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('validates for a correct input with id', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validId
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validId,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('prompts before removing the specified retention event when force option not passed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validId
      })
    });


    assert(promptIssued);
  });

  it('aborts removing the specified retention event when force option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validId
      })
    });
    assert(deleteSpy.notCalled);
  });

  it('Correctly deletes retention event by id', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validId
      })
    });
  });

  it('Correctly deletes retention event by id when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validId,
        force: true
      })
    });
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        code: "generalException",
        message: "Can't remove retention event"
      }
    };

    sinon.stub(request, 'delete').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validId,
        force: true
      })
    }), new CommandError(error.error.message));
  });
});