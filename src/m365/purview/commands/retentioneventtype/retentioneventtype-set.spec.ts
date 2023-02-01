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
const command: Command = require('./retentioneventtype-set');

describe(commands.RETENTIONEVENTTYPE_SET, () => {
  const validId = 'e554d69c-0992-4f9b-8a66-fca3c4d9c531';
  const newDisplayName = 'New display name';
  const description = 'Updated description';

  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.patch
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
    assert.strictEqual(command.name, commands.RETENTIONEVENTTYPE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation with valid id but no other option specified', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with valid id and a single option specified', async () => {
    const actual = await command.validate({ options: { id: validId, newDisplayName: newDisplayName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly sets title and description of a specific retention event type by id', async () => {
    const requestBody = {
      displayName: newDisplayName,
      description: description
    };

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: validId, newDisplayName: newDisplayName, description: description, verbose: true } });
    assert.deepStrictEqual(loggerLogToStderrSpy.lastCall.args[0].data, requestBody);
  });

  it('throws an error when we execute the command using application permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    await assert.rejects(command.action(logger, { options: { id: validId } }),
      new CommandError('This command does not support application permissions.'));
  });

  it('handles error when event type does not exist', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${validId}`) {
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
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: validId
      }
    }), new CommandError(`There is no rule matching identity 'ca0e1f8d-4e42-4a81-be85-022502d70c4f'.`));
  });
});