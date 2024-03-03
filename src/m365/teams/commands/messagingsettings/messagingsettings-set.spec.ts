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
import command from './messagingsettings-set.js';

describe(commands.MESSAGINGSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGINGSETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('sets the allowUserEditMessages setting to true', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.data) === JSON.stringify({
          messagingSettings: {
            allowUserEditMessages: true
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', allowUserEditMessages: true }
    } as any);
  });

  it('sets allowOwnerDeleteMessages, allowTeamMentions and allowChannelMentions to true', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.data) === JSON.stringify({
          messagingSettings: {
            allowOwnerDeleteMessages: true,
            allowTeamMentions: true,
            allowChannelMentions: true
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', allowOwnerDeleteMessages: true, allowTeamMentions: true, allowChannelMentions: true }
    } as any);
  });

  it('should handle Microsoft graph error response', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee`) {
        throw {
          "error": {
            "code": "ItemNotFound",
            "message": "No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee",
            "innerError": {
              "request-id": "27b49647-a335-48f8-9a7c-f1ed9b976aaa",
              "date": "2019-04-05T12:16:48"
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
        allowOwnerDeleteMessages: true,
        allowTeamMentions: true,
        allowChannelMentions: true
      }
    } as any), new CommandError('No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'));
  });

  it('fails validation if the teamId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if allowUserEditMessages is doublicated', async () => {
    const actual = await command.validate({
      options: {
        teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
        allowUserEditMessages: [true, false]
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if allowUserEditMessages is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
        allowUserEditMessages: false
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});