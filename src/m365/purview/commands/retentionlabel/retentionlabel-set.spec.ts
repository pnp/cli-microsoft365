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
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import { session } from '../../../../utils/session.js';
import command, { options } from './retentionlabel-set.js';

describe(commands.RETENTIONLABEL_SET, () => {
  const validId = 'e554d69c-0992-4f9b-8a66-fca3c4d9c531';

  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONLABEL_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with valid id but no other option specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when specifying an invalid value for behaviorDuringRetentionPeriod', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId, behaviorDuringRetentionPeriod: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when specifying an invalid value for actionAfterRetentionPeriod', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId, actionAfterRetentionPeriod: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when specifying an invalid value for retentionTrigger', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId, retentionTrigger: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when specifying an invalid value for defaultRecordBehavior', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId, defaultRecordBehavior: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with valid id and a single option specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId, retentionDuration: '180' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with valid id and multipe options specified', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId, retentionDuration: '180', actionAfterRetentionPeriod: 'none' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: validId, retentionDuration: '180', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('correctly sets field retentionDays and actionAfterRetentionPeriod of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: validId, retentionDuration: '180', actionAfterRetentionPeriod: 'none', verbose: true }) });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('correctly sets field retentionTrigger, defaultRecordBehavior and behaviorDuringRetentionPeriod of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: validId, retentionTrigger: 'dateLabeled', defaultRecordBehavior: 'startLocked', behaviorDuringRetentionPeriod: 'retainAsRecord', verbose: true }) });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('correctly sets field descriptionForUsers, descriptionForAdmins and labelToBeApplied of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: validId, descriptionForUsers: 'description for users', descriptionForAdmins: 'description for admins', labelToBeApplied: 'label to be applied', verbose: true }) });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('fails to set field retentionDuration of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        throw 'Error occurred when updating the retention label';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: validId, retentionDuration: '180', actionAfterRetentionPeriod: 'none' }) }), new CommandError('Error occurred when updating the retention label'));
  });
});