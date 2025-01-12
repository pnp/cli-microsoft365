import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './report-settings-set.js';

describe(commands.REPORT_TENANTSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;



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
    assert.strictEqual(command.name, commands.REPORT_TENANTSETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('logs verbose message when verbose option is enabled', async () => {
    const logToStderrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/reportSettings`) {
        return Promise.resolve();
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { hideUserInformation: true, verbose: true } });

    assert(logToStderrSpy.calledWith('Updating report settings displayConcealedNames to true'));
  });


  it('patches the tenant settings report with the specified settings', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/reportSettings`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, {
      options: { hideUserInformation: true }
    });
  });

  it('handles error when retrieving tenant report settings failed', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/reportSettings`) {
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