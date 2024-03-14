
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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    assert.strictEqual(command.name.startsWith(commands.ENGAGE_COMMUNITY_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/beta/employeeExperience/communities/invalid') {
        throw {
          "error": {
            "code": "badRequest",
            "message": "Bad request.",
            "innerError": {
              "date": "2024-03-14T10:59:12",
              "request-id": "ac728cbd-0978-473b-ab4d-63a51312004a",
              "client-request-id": "6302be7a-9539-138a-26cc-eaa41245ce41"
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(() => command.action(logger, {
      options: {
        id: 'invalid',
        verbose: true
      }
    }), new CommandError(`Bad request.`));
  });

  it('gets community by id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/beta/employeeExperience/communities/eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIxNTU1MjcwOTQyNzIifQ') {
        return Promise.resolve(
          {
            "id": "eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIxNTU1MjcwOTQyNzIifQ",
            "displayName": "New Employee Onboarding",
            "description": "New Employee Onboarding",
            "privacy": "public",
            "groupId": "54dda9b2-2df1-4ce8-ae1e-b400956b5b34"
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIxNTU1MjcwOTQyNzIifQ' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIxNTU1MjcwOTQyNzIifQ');
  });
});