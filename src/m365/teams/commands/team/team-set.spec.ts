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
import command from './team-set.js';

describe(commands.TEAM_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    assert.strictEqual(command.name, commands.TEAM_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('sets the visibility settings correctly', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.data) === JSON.stringify({
          visibility: 'Public'
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', visibility: 'Public' }
    } as any);
  });

  it('sets the mailNickName correctly', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.data) === JSON.stringify({
          mailNickName: 'NewNickName'
        })) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', mailNickName: 'NewNickName' }
    } as any);
  });

  it('sets the description settings correctly', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.data) === JSON.stringify({
          description: 'desc'
        })) {
        return {};
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', description: 'desc' }
    } as any);
  });

  it('sets the classification settings correctly', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.data) === JSON.stringify({
          classification: 'MBI'
        })) {
        return {};
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { debug: true, id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', classification: 'MBI' }
    } as any);
  });

  it('should handle Microsoft graph error response', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
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

    await assert.rejects(command.action(logger, {
      options: {
        id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
        name: 'NewName'
      }
    } as any), new CommandError('No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if visibility is not a valid visibility Private|Public', async () => {
    const actual = await command.validate({
      options: {
        id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
        visibility: 'hidden'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });
});
