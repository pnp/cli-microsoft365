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
const command: Command = require('./guestsettings-set');

describe(commands.GUESTSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GUESTSETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        allowCreateUpdateChannels: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('sets the allowDeleteChannels setting to true', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.data) === JSON.stringify({
          guestSettings: {
            allowDeleteChannels: true
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowDeleteChannels: true }
    } as any);
  });

  it('sets allowCreateUpdateChannels and allowDeleteChannels to true', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.data) === JSON.stringify({
          guestSettings: {
            allowCreateUpdateChannels: true,
            allowDeleteChannels: true
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: true, allowDeleteChannels: true }
    } as any);
  });

  it('correctly handles error when updating guest settings', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };
    sinon.stub(request, 'patch').rejects(error);

    await assert.rejects(command.action(logger, { options: { teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowDeleteChannels: true } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if the teamId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: false
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is true', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: false
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is true', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});