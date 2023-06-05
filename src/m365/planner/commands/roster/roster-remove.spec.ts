import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./roster-remove');

describe(commands.PLAN_REMOVE, () => {
  const validRosterId = 'CRp0hFSovEedkXtcX3WnS5gAGgch';

  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the specified Roster when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        id: validRosterId
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified Roster when confirm option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        id: validRosterId
      }
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
      options: {
        verbose: true,
        id: validRosterId,
        confirm: true
      }
    });
  });

  it('correctly deletes Roster by id when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        id: validRosterId
      }
    });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'delete').rejects({
      error: {
        message: 'The requested item is not found.'
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: validRosterId,
        confirm: true
      }
    }), new CommandError('The requested item is not found.'));
  });
});
