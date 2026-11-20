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
import command, { options } from './funsettings-set.js';

describe(commands.FUNSETTINGS_SET, () => {
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
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FUNSETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets allowGiphy settings to false', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowGiphy: false
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: false })
    });
  });

  it('sets allowGiphy settings to true', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowGiphy: true
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: true })
    });
  });

  it('sets giphyContentRating to moderate', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'moderate'
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'moderate' })
    });
  });

  it('sets giphyContentRating to strict', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'strict'
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'strict' })
    });
  });

  it('sets allowStickersAndMemes to true', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: true
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: true })
    });
  });

  it('sets allowStickersAndMemes to false', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: false
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: false })
    });
  });


  it('sets allowCustomMemes to true', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: true
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: true })
    });
  });

  it('sets allowCustomMemes to false', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: false })
    });
  });

  it('sets allowCustomMemes to false (debug)', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.data) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({ debug: true, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: false })
    });
  });

  it('correctly handles random API error', async () => {
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

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        allowGiphy: true,
        giphyContentRating: "moderate",
        allowStickersAndMemes: false,
        allowCustomMemes: true
      })
    }), new CommandError('An error has occurred'));
  });

  it('fails validation if teamId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when teamId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when giphyContentRating is moderate or strict', () => {
    const actualModerate = commandOptionsSchema.safeParse({ teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'moderate' });
    const actualStrict = commandOptionsSchema.safeParse({ teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'strict' });
    assert.strictEqual(actualModerate.success, true);
    assert.strictEqual(actualStrict.success, true);
  });

  it('fails validation when giphyContentRating is not moderate or strict', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'somethingelse' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when allowStickersAndMemes is a valid boolean', () => {
    const actualTrue = commandOptionsSchema.safeParse({ teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: true });
    const actualFalse = commandOptionsSchema.safeParse({ teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: false });
    assert.strictEqual(actualTrue.success, true);
    assert.strictEqual(actualFalse.success, true);
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