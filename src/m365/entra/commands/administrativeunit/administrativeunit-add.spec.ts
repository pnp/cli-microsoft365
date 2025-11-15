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
import command from './administrativeunit-add.js';
import { options } from './administrativeunit-get.js';

describe(commands.ADMINISTRATIVEUNIT_ADD, () => {
  const administrativeUnitReponse: any = {
    id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    displayName: 'European Division',
    description: null,
    visibility: null
  };

  const administrativeUnitWithDirectoryExtensionReponse: any = {
    id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
    displayName: 'European Division',
    description: null,
    visibility: null,
    extension_b7d8e648520f41d3b9c0fdeb91768a0a_jobGroupTracker: 'JobGroupN'
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

  it('allows unknown options', () => {
    assert.strictEqual(command.allowUnknownOptions(), true);
  });

  it('fails validation when displayName is not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when displayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: 'European Division'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when displayName, description and hiddenMembership are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: 'European Division',
      description: 'European Division Administration',
      hiddenMembership: true
    });
    assert.strictEqual(actual.success, true);
  });

  it('creates an administrative unit with a specific display name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/directory/administrativeUnits') {
        return administrativeUnitReponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'European Division' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: 'European Division',
      description: undefined,
      visibility: null
    });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitReponse));
  });

  it('creates an administrative unit with a specific display name and description', async () => {
    const privateAdministrativeUnitResponse = { ...administrativeUnitReponse };
    privateAdministrativeUnitResponse.description = 'European Division Administration';

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/directory/administrativeUnits') {
        return administrativeUnitReponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'European Division', description: 'European Division Administration' }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: 'European Division',
      description: 'European Division Administration',
      visibility: null
    });
    assert(loggerLogSpy.calledOnceWith(administrativeUnitReponse));
  });

  it('creates an administrative unit with unknown options', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/directory/administrativeUnits') {
        return administrativeUnitWithDirectoryExtensionReponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        displayName: 'European Division',
        description: 'European Division Administration',
        extension_b7d8e648520f41d3b9c0fdeb91768a0a_jobGroupTracker: 'JobGroupN'
      })
    });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: 'European Division',
      description: 'European Division Administration',
      visibility: null,
      extension_b7d8e648520f41d3b9c0fdeb91768a0a_jobGroupTracker: 'JobGroupN'
    });
    assert(loggerLogSpy.calledOnceWith(administrativeUnitWithDirectoryExtensionReponse));
  });

  it('creates a hidden administrative unit with a specific display name and description', async () => {
    const privateAdministrativeUnitResponse = { ...administrativeUnitReponse };
    privateAdministrativeUnitResponse.description = 'European Division Administration';
    privateAdministrativeUnitResponse.visibility = 'HiddenMembership';

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/directory/administrativeUnits') {
        return administrativeUnitReponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'European Division', description: 'European Division Administration', hiddenMembership: true }) });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: 'European Division',
      description: 'European Division Administration',
      visibility: 'HiddenMembership'
    });
    assert(loggerLogSpy.calledOnceWith(administrativeUnitReponse));
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'European Division' }) } as any), new CommandError('Invalid request'));
  });


});