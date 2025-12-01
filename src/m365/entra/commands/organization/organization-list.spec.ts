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
import command, { options } from './organization-list.js';

describe(commands.ORGANIZATION_LIST, () => {
  const response = {
    "id": "e65b162c-6f87-4eb1-a24e-1b37d3504663",
    "deletedDateTime": null,
    "businessPhones": [
      "4258828080"
    ],
    "city": null,
    "country": null,
    "countryLetterCode": "IE",
    "createdDateTime": "2023-02-21T19:56:38Z",
    "defaultUsageLocation": null,
    "displayName": "Contoso",
    "isMultipleDataLocationsForServicesEnabled": null,
    "marketingNotificationEmails": [],
    "onPremisesLastSyncDateTime": null,
    "onPremisesSyncEnabled": null,
    "partnerTenantType": null,
    "postalCode": null,
    "preferredLanguage": "en",
    "securityComplianceNotificationMails": [],
    "securityComplianceNotificationPhones": [],
    "state": null,
    "street": null,
    "technicalNotificationMails": [
      "john.doe@contoso.com"
    ],
    "tenantType": "AAD",
    "directorySizeQuota": {
      "used": 1400,
      "total": 300000
    },
    "assignedPlans": [],
    "onPremisesSyncStatus": [],
    "privacyProfile": {
      "contactEmail": "john.doe@contoso.com",
      "statementUrl": ""
    },
    "provisionedPlans": [],
    "verifiedDomains": [
      {
        "capabilities": "Email, OfficeCommunicationsOnline",
        "isDefault": true,
        "isInitial": true,
        "name": "contoso.onmicrosoft.com",
        "type": "Managed"
      },
      {
        "capabilities": "Email, OfficeCommunicationsOnline, MoeraDomain",
        "isDefault": false,
        "isInitial": false,
        "name": "contoso2.onmicrosoft.com",
        "type": "Managed"
      }
    ]
  };

  let log: any[];
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ORGANIZATION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'tenantType']);
  });

  it('should get a list of organizations', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization`) {
        return {
          value: [
            response
          ]
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly([response]));
  });

  it('should get a list of organizations with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization?$select=id,displayName`) {
        return {
          value: [
            response
          ]
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      properties: 'id,displayName',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(loggerLogSpy.calledOnceWithExactly([response]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        error: {
          code: "accessDenied",
          message: "Request Authorization failed",
          innerError: {
            message: "Request Authorization failed",
            'request-id': "f92d9230-3297-4d2b-9ac5-e7b2abc32d4f",
            'client-request-id': "f92d9230-3297-4d2b-9ac5-e7b2abc32d4f"
          }
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      verbose: true
    });
    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError('Request Authorization failed'));
  });
});