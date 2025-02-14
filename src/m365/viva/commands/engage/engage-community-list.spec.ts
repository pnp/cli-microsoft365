import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './engage-community-list.js';

describe(commands.ENGAGE_COMMUNITY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.ENGAGE_COMMUNITY_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'privacy']);
  });

  it(`should get a list of Viva Engage communities`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities`) {
        return {
          value: [
            {
              "id": "eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9",
              "displayName": "All Company",
              "description": "This is the default group for everyone in the network",
              "privacy": "public",
              "groupId": "7c99afd7-9f3a-49e2-b105-4ee36314350c"
            },
            {
              "id": "eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI1MjM3OTQ0MzIwMSJ9",
              "displayName": "Software Engineers",
              "description": "The group for all developers",
              "privacy": "private",
              "groupId": "c5122932-a925-494d-8b27-f16aa00d41bf"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { verbose: true }
    });

    assert(
      loggerLogSpy.calledWith([
        {
          "id": "eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9",
          "displayName": "All Company",
          "description": "This is the default group for everyone in the network",
          "privacy": "public",
          "groupId": "7c99afd7-9f3a-49e2-b105-4ee36314350c"
        },
        {
          "id": "eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI1MjM3OTQ0MzIwMSJ9",
          "displayName": "Software Engineers",
          "description": "The group for all developers",
          "privacy": "private",
          "groupId": "c5122932-a925-494d-8b27-f16aa00d41bf"
        }
      ])
    );
  });

  it('handles error when retrieving Viva Engage communities failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred')
    );
  });
});