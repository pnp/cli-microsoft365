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
import command from './m365group-renew.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.M365GROUP_RENEW, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const groupId = '28beab62-7540-4db1-a23f-29a6018a3848';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);
    auth.connection.active = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      global.setTimeout
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_RENEW);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('renews expiration of a group specified by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848/renew/') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: groupId, verbose: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('renews expiration of a group specified by displayName', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848/renew/') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Finance', verbose: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('renews expiration the specified group (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848/renew/') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: groupId } });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles error when group is not found', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'The remote server returned an error: (404) Not Found.' } } } });

    await assert.rejects(command.action(logger, { options: { id: groupId } }),
      new CommandError('The remote server returned an error: (404) Not Found.'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { id: groupId } }),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
