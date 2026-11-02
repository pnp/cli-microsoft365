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
import command, { options } from './guestsettings-list.js';

describe(commands.GUESTSETTINGS_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GUESTSETTINGS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists guest settings for a Microsoft Team', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=guestSettings`) {
        return {
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: "2609af39-7775-4f94-a3dc-0dd67657e900"
      })
    });
    assert(loggerLogSpy.calledWith({
      "allowCreateUpdateChannels": false,
      "allowDeleteChannels": false
    }));
  });

  it('lists guest settings for a Microsoft Team (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=guestSettings`) {
        return {
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        teamId: "2609af39-7775-4f94-a3dc-0dd67657e900"
      })
    });
    assert(loggerLogSpy.calledWith({
      "allowCreateUpdateChannels": false,
      "allowDeleteChannels": false
    }));
  });

  it('correctly handles error when listing guest settings for a Microsoft Team', async () => {
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
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: "2609af39-7775-4f94-a3dc-0dd67657e900"
      })
    }), new CommandError('An error has occurred'));
  });

  it('fails validation if teamId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when teamId is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ teamId: '2609af39-7775-4f94-a3dc-0dd67657e900' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '26be5f98-e66b-4e0a-bc37-1e6b1b8e5b7b',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('lists all properties for output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=guestSettings`) {
        return {
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: "2609af39-7775-4f94-a3dc-0dd67657e900",
        output: 'json'
      })
    });
    assert(loggerLogSpy.calledWith({
      "allowCreateUpdateChannels": false,
      "allowDeleteChannels": false
    }));
  });
});
