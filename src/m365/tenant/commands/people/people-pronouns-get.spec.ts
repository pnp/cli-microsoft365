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
import command from './people-pronouns-get.js';

describe(commands.PEOPLE_PRONOUNS_GET, () => {
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
    assert.strictEqual(command.name, commands.PEOPLE_PRONOUNS_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`should get pronouns settings`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/pronouns`) {
        return {
          isEnabledInOrganization: true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { verbose: true }
    });

    assert(
      loggerLogSpy.calledOnceWithExactly({
        isEnabledInOrganization: true
      })
    );
  });

  it('handles error when retrieving pronouns settings failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/pronouns`) {
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