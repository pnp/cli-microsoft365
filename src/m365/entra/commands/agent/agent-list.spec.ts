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
import command, { options } from './agent-list.js';

describe(commands.AGENT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const response = [
    {
      "id": "819f4e7f-a913-448a-9b5c-2cae064cabe7",
      "deletedDateTime": null,
      "accountEnabled": true,
      "ageGroup": null,
      "businessPhones": [],
      "city": null,
      "companyName": null,
      "consentProvidedForMinor": null,
      "country": null,
      "createdDateTime": "2026-07-14T04:19:25Z",
      "creationType": null,
      "department": null,
      "displayName": "Sales Agent",
      "employeeId": null,
      "employeeHireDate": null,
      "employeeLeaveDateTime": null,
      "employeeType": null,
      "externalUserState": null,
      "externalUserStateChangeDateTime": null,
      "faxNumber": null,
      "givenName": null,
      "isLicenseReconciliationNeeded": false,
      "jobTitle": null,
      "legalAgeGroupClassification": null,
      "mail": null,
      "mailNickname": "SalesAgent",
      "mobilePhone": null,
      "onPremisesDistinguishedName": null,
      "onPremisesDomainName": null,
      "onPremisesImmutableId": null,
      "onPremisesLastSyncDateTime": null,
      "onPremisesSecurityIdentifier": null,
      "onPremisesSamAccountName": null,
      "onPremisesSyncEnabled": null,
      "onPremisesUserPrincipalName": null,
      "otherMails": [],
      "passwordPolicies": null,
      "officeLocation": null,
      "postalCode": null,
      "preferredDataLocation": null,
      "preferredLanguage": null,
      "proxyAddresses": [],
      "refreshTokensValidFromDateTime": null,
      "imAddresses": [],
      "isResourceAccount": null,
      "showInAddressList": null,
      "securityIdentifier": "S-1-12-1-2174701183-1149937939-2922142875-3869985798",
      "signInSessionsValidFromDateTime": null,
      "state": null,
      "streetAddress": null,
      "surname": null,
      "usageLocation": null,
      "userPrincipalName": "SalesAgent@contoso.com",
      "externalUserConvertedOn": null,
      "userType": "Member",
      "identityParentId": "54d8c728-decb-477b-beb6-19570d8a51ab",
      "agentIdentityBlueprintId": "822c92e8-29d7-4f83-903a-744d00628003",
      "employeeOrgData": null,
      "passwordProfile": null,
      "identityParent": {
        "id": "54d8c728-decb-477b-beb6-19570d8a51ab"
      },
      "assignedLicenses": [],
      "assignedPlans": [],
      "authorizationInfo": {
        "certificateUserIds": []
      },
      "identities": [
        {
          "signInType": "userPrincipalName",
          "issuer": "contoso.com",
          "issuerAssignedId": "SalesAgent@contoso.com"
        }
      ],
      "onPremisesProvisioningErrors": [],
      "onPremisesExtensionAttributes": {
        "extensionAttribute1": null,
        "extensionAttribute2": null,
        "extensionAttribute3": null,
        "extensionAttribute4": null,
        "extensionAttribute5": null,
        "extensionAttribute6": null,
        "extensionAttribute7": null,
        "extensionAttribute8": null,
        "extensionAttribute9": null,
        "extensionAttribute10": null,
        "extensionAttribute11": null,
        "extensionAttribute12": null,
        "extensionAttribute13": null,
        "extensionAttribute14": null,
        "extensionAttribute15": null
      },
      "provisionedPlans": [],
      "serviceProvisioningErrors": []
    }
  ];

  const propertiesResponse = [
    {
      "displayName": "Sales Agent",
      "id": "819f4e7f-a913-448a-9b5c-2cae064cabe7"
    }
  ];

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
    assert.strictEqual(command.name, commands.AGENT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName']);
  });

  it(`should get a list of agents`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/microsoft.graph.agentUser`) {
        return {
          value: response
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledWith(response));
  });

  it(`should get a list of agents with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/microsoft.graph.agentUser?$select=id,displayName`) {
        return {
          value: propertiesResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ properties: 'id,displayName' }) });
    assert(loggerLogSpy.calledWith(propertiesResponse));
  });

  it('handles error when retrieving agents list failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/microsoft.graph.agentUser`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('An error has occurred')
    );
  });
});