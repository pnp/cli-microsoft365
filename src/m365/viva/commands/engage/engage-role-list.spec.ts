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
import command from './engage-role-list.js';

describe(commands.ENGAGE_ROLE_LIST, () => {
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
    assert.strictEqual(command.name, commands.ENGAGE_ROLE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName']);
  });

  it(`should get a list of Viva Engage roles`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles`) {
        return {
          "value": [
            {
              "id": "ec759127-089f-4f91-8dfc-03a30b51cb38",
              "displayName": "Network Admin"
            },
            {
              "id": "966b8ec4-6457-4f22-bd3c-5a2520e98f4a",
              "displayName": "Verified Admin"
            },
            {
              "id": "77aa47ad-96fe-4ecc-8024-fd1ac5e28f17",
              "displayName": "Corporate Communicator"
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
      loggerLogSpy.calledOnceWith([
        {
          "id": "ec759127-089f-4f91-8dfc-03a30b51cb38",
          "displayName": "Network Admin"
        },
        {
          "id": "966b8ec4-6457-4f22-bd3c-5a2520e98f4a",
          "displayName": "Verified Admin"
        },
        {
          "id": "77aa47ad-96fe-4ecc-8024-fd1ac5e28f17",
          "displayName": "Corporate Communicator"
        }
      ])
    );
  });

  it('handles error when retrieving Viva Engage roles failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: {} }),
      new CommandError('An error has occurred')
    );
  });
});