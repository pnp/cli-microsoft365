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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { accessToken } from '../../../../utils/accessToken';
const command: Command = require('./retentioneventtype-remove');

describe(commands.RETENTIONEVENTTYPE_REMOVE, () => {
  const validId = 'e554d69c-0992-4f9b-8a66-fca3c4d9c531';

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
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
      accessToken.isAppOnlyAccessToken,
      Cli.prompt,
      request.delete
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
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { id: validId } });
  });

  it('correctly deletes retention event type by id when prompt confirmed', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: validId, confirm: true } });
  });

  it('throws an error when we execute the command using application permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    await assert.rejects(command.action(logger, { options: { id: validId } }),
      new CommandError('This command does not support application permissions.'));
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
        confirm: true
      }
    }), new CommandError(`There is no rule matching identity 'ca0e1f8d-4e42-4a81-be85-022502d70c4f'.`));
  });
});
