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
import command from './retentionlabel-set.js';
import { session } from '../../../../utils/session.js';

describe(commands.RETENTIONLABEL_SET, () => {
  const validId = 'e554d69c-0992-4f9b-8a66-fca3c4d9c531';

  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONLABEL_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation with valid id but no other option specified', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when specifying an invalid value for behaviorDuringRetentionPeriod', async () => {
    const actual = await command.validate({ options: { id: validId, behaviorDuringRetentionPeriod: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when specifying an invalid value for actionAfterRetentionPeriod', async () => {
    const actual = await command.validate({ options: { id: validId, actionAfterRetentionPeriod: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when specifying an invalid value for retentionTrigger', async () => {
    const actual = await command.validate({ options: { id: validId, retentionTrigger: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when specifying an invalid value for defaultRecordBehavior', async () => {
    const actual = await command.validate({ options: { id: validId, defaultRecordBehavior: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with valid id and a single option specified', async () => {
    const actual = await command.validate({ options: { id: validId, retentionDuration: 180 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with valid id and multipe options specified', async () => {
    const actual = await command.validate({ options: { id: validId, retentionDuration: 180, actionAfterRetentionPeriod: 'none' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly sets field retentionDays and actionAfterRetentionPeriod of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: validId, retentionDuration: 180, actionAfterRetentionPeriod: 'none', verbose: true } });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('correctly sets field retentionTrigger, defaultRecordBehavior and behaviorDuringRetentionPeriod of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: validId, retentionTrigger: 'dateLabeled', defaultRecordBehavior: 'startLocked', behaviorDuringRetentionPeriod: 'retainAsRecord', verbose: true } });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('correctly sets field descriptionForUsers, descriptionForAdmins and labelToBeApplied of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { id: validId, descriptionForUsers: 'description for users', descriptionForAdmins: 'description for admins', labelToBeApplied: 'label to be applied', verbose: true } });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('fails to set field retentionDays of a specific retention label by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        throw 'Error occurred when updating the retention label';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: validId, retentionDays: 180, actionAfterRetentionPeriod: 'none' } }), new CommandError('Error occurred when updating the retention label'));
  });
});