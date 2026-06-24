import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { cli } from '../../../../cli/cli.js';
import command, { options } from './administrativeunit-member-get.js';
import request from '../../../../request.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';

describe(commands.ADMINISTRATIVEUNIT_MEMBER_GET, () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const administrativeUnitName = 'European Division';
  const userId = '64131a70-beb9-4ccb-b590-4401e58446ec';

  const userTransformedResponse = {
    "id": "64131a70-beb9-4ccb-b590-4401e58446ec",
    "businessPhones": [
      "+20 255501070"
    ],
    "displayName": "Pradeep Gupta",
    "givenName": "Pradeep",
    "jobTitle": "Accountant",
    "mail": "PradeepG@4wrvkx.onmicrosoft.com",
    "mobilePhone": null,
    "officeLocation": "98/2202",
    "preferredLanguage": "en-US",
    "surname": "Gupta",
    "userPrincipalName": "PradeepG@4wrvkx.onmicrosoft.com",
    "type": "user"
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_MEMBER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when member id and administrativeUnitId are GUIDs', () => {
    const actual = commandOptionsSchema.safeParse({ id: userId, administrativeUnitId: administrativeUnitId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if member id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: 'invalid', administrativeUnitId: administrativeUnitId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if administrativeUnitId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: userId, administrativeUnitId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both administrativeUnitId and administrativeUnitName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ id: userId, administrativeUnitId: administrativeUnitId, administrativeUnitName: administrativeUnitName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both administrativeUnitId and administrativeUnitName options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ id: userId });
    assert.strictEqual(actual.success, false);
  });

  it('get member info for an administrative unit specified by id and member specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}`) {
        return {
          "@odata.type": "#microsoft.graph.user",
          "id": "64131a70-beb9-4ccb-b590-4401e58446ec",
          "businessPhones": [
            "+20 255501070"
          ],
          "displayName": "Pradeep Gupta",
          "givenName": "Pradeep",
          "jobTitle": "Accountant",
          "mail": "PradeepG@4wrvkx.onmicrosoft.com",
          "mobilePhone": null,
          "officeLocation": "98/2202",
          "preferredLanguage": "en-US",
          "surname": "Gupta",
          "userPrincipalName": "PradeepG@4wrvkx.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: userId, administrativeUnitId: administrativeUnitId } });

    assert(loggerLogSpy.calledOnceWithExactly(userTransformedResponse));
  });

  it('get member info for an administrative unit specified by name and member specified by id (verbose)', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}`) {
        return {
          "@odata.type": "#microsoft.graph.user",
          "id": "64131a70-beb9-4ccb-b590-4401e58446ec",
          "businessPhones": [
            "+20 255501070"
          ],
          "displayName": "Pradeep Gupta",
          "givenName": "Pradeep",
          "jobTitle": "Accountant",
          "mail": "PradeepG@4wrvkx.onmicrosoft.com",
          "mobilePhone": null,
          "officeLocation": "98/2202",
          "preferredLanguage": "en-US",
          "surname": "Gupta",
          "userPrincipalName": "PradeepG@4wrvkx.onmicrosoft.com"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: userId, administrativeUnitName: administrativeUnitName, verbose: true } });

    assert(loggerLogSpy.calledOnceWithExactly(userTransformedResponse));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { id: userId, administrativeUnitId: administrativeUnitId } }), new CommandError(errorMessage));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: userId, administrativeUnitId: administrativeUnitId } } as any), new CommandError('Invalid request'));
  });

  it('retrieves selected properties of a member for an administrative unit specified by id and member specified by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${userId}?$select=id,displayName&$expand=manager($select=displayName),drive($select=id)`) {
        return {
          "@odata.type": "#microsoft.graph.user",
          "id": "64131a70-beb9-4ccb-b590-4401e58446ec",
          "displayName": "Pradeep Gupta",
          "manager": {
            "displayName": "Adele Vance"
          },
          "drive": {
            "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT"
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: userId, administrativeUnitId: administrativeUnitId, properties: 'id,displayName,manager/displayName,drive/id' } });

    assert(loggerLogSpy.calledOnceWithExactly({
      "id": "64131a70-beb9-4ccb-b590-4401e58446ec",
      "displayName": "Pradeep Gupta",
      "manager": {
        "displayName": "Adele Vance"
      },
      "drive": {
        "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT"
      },
      "type": "user"
    }));
  });
});