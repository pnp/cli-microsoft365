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
import command, { options } from './team-archive.js';

describe(commands.TEAM_ARCHIVE, () => {
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
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TEAM_ARCHIVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid guid.', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when the input is correct', () => {
    const actual = commandOptionsSchema.safeParse({ id: '15d7a78e-fd77-4599-97a5-dbb6372846c5' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when no option is specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and name are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: '15d7a78e-fd77-4599-97a5-dbb6372846c5', name: 'Finance' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: '15d7a78e-fd77-4599-97a5-dbb6372846c5', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('fails when team name does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq 'Finance'&$select=id`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 0,
          "value": []
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        name: 'Finance'
      })
    }), new CommandError("The specified team 'Finance' does not exist."));
  });

  it('archives a Microsoft Team by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
      })
    });
  });

  it('archives a Microsoft Team by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq 'Finance'&$select=id`) {
        return {
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000"
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/archive`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        name: 'Finance'
      })
    });
  });

  it('archives a Microsoft Teams teams when \'shouldSetSpoSiteReadOnlyForMembers\' specified', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6',
        shouldSetSpoSiteReadOnlyForMembers: true
      })
    });
    assert.strictEqual(postStub.lastCall.args[0].data.shouldSetSpoSiteReadOnlyForMembers, true);
  });

  it('should set to false value when \'shouldSetSpoSiteReadOnlyForMembers\' not specified', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
      })
    });
    assert.strictEqual(postStub.lastCall.args[0].data["shouldSetSpoSiteReadOnlyForMembers"], false);
  });

  it('should correctly handle graph error response', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        throw {
          "error": {
            "code": "ItemNotFound",
            "message": "No team found with Group Id f5dba91d-6494-4d5e-89a7-ad832f6946d6",
            "innerError": {
              "request-id": "ad0c0a4f-a4fc-4567-8ae1-1150db48b620",
              "date": "2019-04-05T15:51:43"
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6' }) }), new CommandError('No team found with Group Id f5dba91d-6494-4d5e-89a7-ad832f6946d6'));
  });

  it('correctly handles error when archiving a team', async () => {
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
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
      })
    }), new CommandError('An error has occurred'));
  });
});
