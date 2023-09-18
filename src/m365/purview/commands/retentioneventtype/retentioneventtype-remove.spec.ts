import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './retentioneventtype-remove.js';

describe(commands.RETENTIONEVENTTYPE_REMOVE, () => {
  const validId = 'e554d69c-0992-4f9b-8a66-fca3c4d9c531';

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
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    promptOptions = undefined;
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.promptForConfirmation,
      request.delete
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENTTYPE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified retention event type when confirm option not passed', async () => {
    await command.action(logger, { options: { id: validId } });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified retention event type when confirm option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, { options: { id: validId } });
    assert(deleteSpy.notCalled);
  });

  it('correctly deletes retention event type by id', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { id: validId } });
  });

  it('correctly deletes retention event type by id when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: validId, force: true } });
  });

  it('handles error when retention event type does not exist', async () => {
    sinon.stub(request, 'delete').callsFake(async () => {
      throw {
        'error': {
          'code': 'UnknownError',
          'message': `There is no rule matching identity 'ca0e1f8d-4e42-4a81-be85-022502d70c4f'.`,
          'innerError': {
            'date': '2023-01-31T21:51:20',
            'request-id': '8160d45b-55b3-4f2a-b741-1da41c454809',
            'client-request-id': '8160d45b-55b3-4f2a-b741-1da41c454809'
          }
        }
      };
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: validId,
        force: true
      }
    }), new CommandError(`There is no rule matching identity 'ca0e1f8d-4e42-4a81-be85-022502d70c4f'.`));
  });
});
