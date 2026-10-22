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
import command, { options } from './roster-remove.js';

describe(commands.ROSTER_REMOVE, () => {
  const validRosterId = 'CRp0hFSovEedkXtcX3WnS5gAGgch';

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
    assert.strictEqual(command.name, commands.ROSTER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with valid id', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validRosterId
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validRosterId,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('prompts before removing the specified Roster when force option not passed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validRosterId
      })
    });

    assert(promptIssued);
  });

  it('aborts removing the specified Roster when force option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validRosterId
      })
    });
    assert(deleteSpy.notCalled);
  });

  it('correctly deletes Roster by id', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        id: validRosterId,
        force: true
      })
    });
  });

  it('correctly deletes Roster by id when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validRosterId
      })
    });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'delete').rejects({
      error: {
        message: 'The requested item is not found.'
      }
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validRosterId,
        force: true
      })
    }), new CommandError('The requested item is not found.'));
  });
});
