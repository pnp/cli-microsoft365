import assert from 'assert';
import sinon from "sinon";
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from "../../../../cli/Logger.js";
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from "../../commands.js";
import command, { options } from './administrativeunit-get.js';

describe(commands.ADMINISTRATIVEUNIT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const administrativeUnitsReponse = {
    value: [
      {
        id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
        displayName: 'European Division',
        visibility: 'HiddenMembership'
      },
      {
        id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
        displayName: 'Asian Division',
        visibility: null
      }
    ]
  };
  const validId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const validDisplayName = 'European Division';

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when id or displayName are not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when id is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validId
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when displayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: validDisplayName
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '123'
    });
    assert.strictEqual(actual.success, false);
  });

  it('retrieves information about the specified administrative unit by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${validId}`) {
        return administrativeUnitsReponse.value[0];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: validId }) });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitsReponse.value[0]));
  });

  it('retrieves information about the specified administrative unit by id with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${validId}?$select=id,displayName,visibility`) {
        return administrativeUnitsReponse.value[0];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: validId, properties: 'id,displayName,visibility' }) });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitsReponse.value[0]));
  });

  it('retrieves information about the specified administrative unit by displayName', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').resolves(administrativeUnitsReponse.value[0]);

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: validDisplayName }) });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitsReponse.value[0]));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: validId }) }), new CommandError(errorMessage));
  });
});