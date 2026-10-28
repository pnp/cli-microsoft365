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
import command, { options } from './team-remove.js';

describe(commands.TEAM_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
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

    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TEAM_REMOVE);
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

  it('fails validation when both id and name are specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: '15d7a78e-fd77-4599-97a5-dbb6372846c5', name: 'Finance' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when neither id nor name is specified', () => {
    const actual = commandOptionsSchema.safeParse({});
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
        name: 'Finance',
        force: true
      })
    }), new CommandError("The specified team 'Finance' does not exist."));
  });

  it('prompts before removing the specified team by id when force option not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ id: "00000000-0000-0000-0000-000000000000" }) });
    assert(promptIssued);
  });

  it('prompts before removing the specified team by name when force option not passed (debug)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: "Finance" }) });
    assert(promptIssued);
  });

  it('aborts removing the specified team when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    await command.action(logger, { options: commandOptionsSchema.parse({ id: "00000000-0000-0000-0000-000000000000" }) });
    assert(postSpy.notCalled);
  });

  it('aborts removing the specified team when force option not passed and prompt not confirmed (debug)', async () => {
    const postSpy = sinon.spy(request, 'delete');
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: "00000000-0000-0000-0000-000000000000" }) });
    assert(postSpy.notCalled);
  });

  it('removes the specified team when prompt confirmed (debug)', async () => {
    let teamsDeleteCallIssued = false;

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000`) {
        teamsDeleteCallIssued = true;
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: "00000000-0000-0000-0000-000000000000" }) });
    assert(teamsDeleteCallIssued);
  });

  it('removes the specified team by id without prompting when confirmed specified', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: "00000000-0000-0000-0000-000000000000", force: true }) });
  });

  it('removes the specified team by name without prompting when confirmed specified', async () => {
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

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: "Finance", force: true }) });
  });

  it('should handle Microsoft graph error response', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee`) {
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

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }) }), new CommandError('No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'));
  });
});
