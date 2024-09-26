import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { misc } from '../../../../utils/misc.js';
import { MockRequests } from '../../../../utils/MockRequest.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';
import command from './administrativeunit-add.js';

const administrativeUnitResponse: any = {
  id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
  displayName: 'European Division',
  description: null,
  visibility: null
};

export const mocks = {
  administrativeUnits: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/directory/administrativeUnits',
      method: 'POST'
    },
    response: {
      body: administrativeUnitResponse
    }
  }
} satisfies MockRequests;

describe(commands.ADMINISTRATIVEUNIT_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.ADMINISTRATIVEUNIT_ADD]);
  });

  it('creates an administrative unit with a specific display name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.administrativeUnits.request.url) {
        return misc.deepClone(mocks.administrativeUnits.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'European Division' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: 'European Division',
      description: undefined,
      visibility: null
    });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitResponse));
  });

  it('creates an administrative unit with a specific display name and description', async () => {
    const privateAdministrativeUnitResponse = { ...misc.deepClone(mocks.administrativeUnits.response.body) };
    privateAdministrativeUnitResponse.description = 'European Division Administration';

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.administrativeUnits.request.url) {
        return privateAdministrativeUnitResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'European Division', description: 'European Division Administration' } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: 'European Division',
      description: 'European Division Administration',
      visibility: null
    });
    assert(loggerLogSpy.calledOnceWith(privateAdministrativeUnitResponse));
  });

  it('creates a hidden administrative unit with a specific display name and description', async () => {
    const privateAdministrativeUnitResponse = { ...misc.deepClone(mocks.administrativeUnits.response.body) };
    privateAdministrativeUnitResponse.description = 'European Division Administration';
    privateAdministrativeUnitResponse.visibility = 'HiddenMembership';

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.administrativeUnits.request.url) {
        return privateAdministrativeUnitResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'European Division', description: 'European Division Administration', hiddenMembership: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: 'European Division',
      description: 'European Division Administration',
      visibility: 'HiddenMembership'
    });
    assert(loggerLogSpy.calledOnceWith(privateAdministrativeUnitResponse));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Invalid request'));
  });

  it('passes validation when only displayName is specified', async () => {
    const actual = await command.validate({ options: { displayName: 'European Division' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the displayName, description and hiddenMembership are specified', async () => {
    const actual = await command.validate({ options: { displayName: 'European Division', description: 'European Division Administration', hiddenMembership: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying displayName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying hiddenMembership', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--hiddenMembership') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});