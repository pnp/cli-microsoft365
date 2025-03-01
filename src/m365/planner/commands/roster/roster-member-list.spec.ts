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
import command from './roster-member-list.js';

describe(commands.ROSTER_MEMBER_LIST, () => {
  const rosterMemberResponse = {
    value: [
      {
        id: "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
        userId: "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
        tenantId: "0cac6cda-2e04-4a3d-9c16-9c91470d7022",
        roles: []
      },
      {
        id: "eb77fbcf-6fe8-458b-985d-1747284793bc",
        userId: "eb77fbcf-6fe8-458b-985d-1747284793bc",
        tenantId: "0cac6cda-2e04-4a3d-9c16-9c91470d7022",
        roles: []
      }
    ]
  };
  const validRosterId = "iryDKm9VLku2HIoC2G-TX5gABJw0";

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
    assert.strictEqual(command.name, commands.ROSTER_MEMBER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves members from a roster', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/members`)) {
        return rosterMemberResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { rosterId: validRosterId, verbose: true } });
    assert(loggerLogSpy.calledWith(rosterMemberResponse.value));
  });


  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The requested item is not found.'
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        rosterId: validRosterId
      }
    }), new CommandError('The requested item is not found.'));
  });
});