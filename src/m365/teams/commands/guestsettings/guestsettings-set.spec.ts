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
import command, { options } from './guestsettings-set.js';

describe(commands.GUESTSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.GUESTSETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowDeleteChannels: true })
    });
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
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: true, allowDeleteChannels: true })
    });
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowDeleteChannels: true }) }), new CommandError('An error has occurred'));
  });

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the teamId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if allowDeleteChannels is false', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
      allowDeleteChannels: false
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if allowDeleteChannels is true', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
      allowDeleteChannels: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if allowCreateUpdateChannels is false', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
      allowCreateUpdateChannels: false
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if allowCreateUpdateChannels is true', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
      allowCreateUpdateChannels: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '26be5f98-e66b-4e0a-bc37-1e6b1b8e5b7b',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });
});