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
import command, { options } from './tenant-settings-set.js';

describe(commands.TENANT_SETTINGS_SET, () => {
  const successResponse = {
    id: '1',
    isPlannerAllowed: true,
    allowCalendarSharing: true,
    allowTenantMoveWithDataLoss: false,
    allowTenantMoveWithDataMigration: false,
    allowRosterCreation: true,
    allowPlannerMobilePushNotifications: true
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      isPlannerAllowed: true,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when no options specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when valid options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      isPlannerAllowed: true,
      allowCalendarSharing: false,
      allowPlannerMobilePushNotifications: false
    });
    assert.strictEqual(actual.success, true);
  });

  it('successfully updates tenant planner settings', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://tasks.office.com/taskAPI/tenantAdminSettings/Settings') {
        return successResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        isPlannerAllowed: true
      })
    });
    assert(loggerLogSpy.calledWith(successResponse));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://tasks.office.com/taskAPI/tenantAdminSettings/Settings') {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        isPlannerAllowed: true
      })
    }), new CommandError('An error has occurred'));
  });
});
