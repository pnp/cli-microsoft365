
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
import command from './engage-community-get.js';

describe(commands.ENGAGE_COMMUNITY_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    assert.strictEqual(command.name, commands.ENGAGE_COMMUNITY_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    const errorMessage = 'Bad request.';
    sinon.stub(request, 'get').rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { id: 'invalid', verbose: true } } as any),
      new CommandError(errorMessage));
  });

  it('gets community by id', async () => {
    const communityId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIxNTU1MjcwOTQyNzIifQ';
    const response = {
      id: communityId,
      displayName: "New Employee Onboarding",
      description: "New Employee Onboarding",
      privacy: "public",
      groupId: "54dda9b2-2df1-4ce8-ae1e-b400956b5b34"
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/communities/${communityId}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: communityId } } as any);
    assert.deepStrictEqual(loggerLogSpy.lastCall.args[0], response);
  });
});