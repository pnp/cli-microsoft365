import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import command from './administrativeunit-member-list.js';
import { settingsNames } from '../../../../settingsNames.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Cli } from '../../../../cli/Cli.js';
import { aadAdministrativeUnit } from '../../../../utils/aadAdministrativeUnit.js';

describe(commands.ADMINISTRATIVEUNIT_MEMBER_LIST, () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const administrativeUnitName = 'European Division';

  const userResponseWithoutMetadata = {
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
  const limitedUserResponseWithoutMetadata = {
    "id": "64131a70-beb9-4ccb-b590-4401e58446ec",
    "displayName": "Pradeep Gupta",
    "manager": {
      "displayName": "Adele Vance"
    },
    "drive": {
      "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT"
    }
  };
  
  const groupResponseWithoutMetadata = {
    "id": "c121c70b-deb1-43f7-8298-9111bf3036b4",
    "deletedDateTime": null,
    "classification": null,
    "createdDateTime": "2023-02-22T06:39:13Z",
    "creationOptions": [
      "Team"
    ],
    "description": "Welcome to the team that we have assembled to create the Mark 8.",
    "displayName": "Mark 8 Project Team",
    "expirationDateTime": null,
    "groupTypes": [
      "Unified"
    ],
    "isAssignableToRole": null,
    "mail": "Mark8ProjectTeam@4wrvkx.onmicrosoft.com",
    "mailEnabled": true,
    "mailNickname": "Mark8ProjectTeam",
    "membershipRule": null,
    "membershipRuleProcessingState": null,
    "onPremisesDomainName": null,
    "onPremisesLastSyncDateTime": null,
    "onPremisesNetBiosName": null,
    "onPremisesSamAccountName": null,
    "onPremisesSecurityIdentifier": null,
    "onPremisesSyncEnabled": null,
    "preferredDataLocation": null,
    "preferredLanguage": null,
    "proxyAddresses": [],
    "renewedDateTime": "2023-02-22T06:39:13Z",
    "resourceBehaviorOptions": [
      "HideGroupInOutlook",
      "SubscribeMembersToCalendarEventsDisabled",
      "WelcomeEmailDisabled"
    ],
    "resourceProvisioningOptions": [
      "Team"
    ],
    "securityEnabled": false,
    "securityIdentifier": "S-1-12-1-3240216331-1140317873-294754434-3023450303",
    "theme": null,
    "visibility": "Public",
    "onPremisesProvisioningErrors": [],
    "serviceProvisioningErrors": []
  };
  const groupTransformedResponse = {
    "id": "c121c70b-deb1-43f7-8298-9111bf3036b4",
    "deletedDateTime": null,
    "classification": null,
    "createdDateTime": "2023-02-22T06:39:13Z",
    "creationOptions": [
      "Team"
    ],
    "description": "Welcome to the team that we have assembled to create the Mark 8.",
    "displayName": "Mark 8 Project Team",
    "expirationDateTime": null,
    "groupTypes": [
      "Unified"
    ],
    "isAssignableToRole": null,
    "mail": "Mark8ProjectTeam@4wrvkx.onmicrosoft.com",
    "mailEnabled": true,
    "mailNickname": "Mark8ProjectTeam",
    "membershipRule": null,
    "membershipRuleProcessingState": null,
    "onPremisesDomainName": null,
    "onPremisesLastSyncDateTime": null,
    "onPremisesNetBiosName": null,
    "onPremisesSamAccountName": null,
    "onPremisesSecurityIdentifier": null,
    "onPremisesSyncEnabled": null,
    "preferredDataLocation": null,
    "preferredLanguage": null,
    "proxyAddresses": [],
    "renewedDateTime": "2023-02-22T06:39:13Z",
    "resourceBehaviorOptions": [
      "HideGroupInOutlook",
      "SubscribeMembersToCalendarEventsDisabled",
      "WelcomeEmailDisabled"
    ],
    "resourceProvisioningOptions": [
      "Team"
    ],
    "securityEnabled": false,
    "securityIdentifier": "S-1-12-1-3240216331-1140317873-294754434-3023450303",
    "theme": null,
    "visibility": "Public",
    "onPremisesProvisioningErrors": [],
    "serviceProvisioningErrors": [],
    "type": "group"
  };

  const deviceResponseWithoutMetadata = {
    "id": "3f9fd7c3-73ad-4ce3-b053-76bb8252964d",
    "deletedDateTime": null,
    "accountEnabled": true,
    "approximateLastSignInDateTime": null,
    "complianceExpirationDateTime": null,
    "createdDateTime": "2023-11-06T06:18:26Z",
    "deviceCategory": null,
    "deviceId": "4c299165-6e8f-4b45-a5ba-c5d250a707ff",
    "deviceMetadata": null,
    "deviceOwnership": null,
    "deviceVersion": null,
    "displayName": "AdeleVence-PC",
    "domainName": null,
    "enrollmentProfileName": null,
    "enrollmentType": null,
    "externalSourceName": null,
    "isCompliant": null,
    "isManaged": null,
    "isRooted": null,
    "managementType": null,
    "manufacturer": null,
    "mdmAppId": null,
    "model": null,
    "onPremisesLastSyncDateTime": null,
    "onPremisesSyncEnabled": null,
    "operatingSystem": "windows",
    "operatingSystemVersion": "10",
    "physicalIds": [],
    "profileType": null,
    "registrationDateTime": null,
    "sourceType": null,
    "systemLabels": [],
    "trustType": null,
    "extensionAttributes": {
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
    "alternativeSecurityIds": [
      {
        "type": 2,
        "identityProvider": null,
        "key": "Y3YxN2E1MWFlYw=="
      }
    ]
  };
  const deviceTransformedResponse = {
    "id": "3f9fd7c3-73ad-4ce3-b053-76bb8252964d",
    "deletedDateTime": null,
    "accountEnabled": true,
    "approximateLastSignInDateTime": null,
    "complianceExpirationDateTime": null,
    "createdDateTime": "2023-11-06T06:18:26Z",
    "deviceCategory": null,
    "deviceId": "4c299165-6e8f-4b45-a5ba-c5d250a707ff",
    "deviceMetadata": null,
    "deviceOwnership": null,
    "deviceVersion": null,
    "displayName": "AdeleVence-PC",
    "domainName": null,
    "enrollmentProfileName": null,
    "enrollmentType": null,
    "externalSourceName": null,
    "isCompliant": null,
    "isManaged": null,
    "isRooted": null,
    "managementType": null,
    "manufacturer": null,
    "mdmAppId": null,
    "model": null,
    "onPremisesLastSyncDateTime": null,
    "onPremisesSyncEnabled": null,
    "operatingSystem": "windows",
    "operatingSystemVersion": "10",
    "physicalIds": [],
    "profileType": null,
    "registrationDateTime": null,
    "sourceType": null,
    "systemLabels": [],
    "trustType": null,
    "extensionAttributes": {
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
    "alternativeSecurityIds": [
      {
        "type": 2,
        "identityProvider": null,
        "key": "Y3YxN2E1MWFlYw=="
      }
    ],
    "type": "device"
  };

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      aadAdministrativeUnit.getAdministrativeUnitByDisplayName,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_MEMBER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when administrativeUnitId is a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if administrativeUnitId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both administrativeUnitId and administrativeUnitName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, administrativeUnitName: administrativeUnitName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both administrativeUnitId and administrativeUnitName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if wrong type is specified', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, type: 'application' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if user type is specified', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, type: 'user' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if group type is specified', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, type: 'group' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if device type is specified', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, type: 'device' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if filter is specified but type is missing', async () => {
    const actual = await command.validate({ options: { administrativeUnitId: administrativeUnitId, filter: "userType eq 'Memmber'" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should get all members of an administrative unit specified by its id when type not specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members`) {
        return {
          value: [
            {
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
            },
            {
              "@odata.type": "#microsoft.graph.group",
              "id": "c121c70b-deb1-43f7-8298-9111bf3036b4",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2023-02-22T06:39:13Z",
              "creationOptions": [
                "Team"
              ],
              "description": "Welcome to the team that we have assembled to create the Mark 8.",
              "displayName": "Mark 8 Project Team",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "Mark8ProjectTeam@4wrvkx.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "Mark8ProjectTeam",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [],
              "renewedDateTime": "2023-02-22T06:39:13Z",
              "resourceBehaviorOptions": [
                "HideGroupInOutlook",
                "SubscribeMembersToCalendarEventsDisabled",
                "WelcomeEmailDisabled"
              ],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-3240216331-1140317873-294754434-3023450303",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": [],
              "serviceProvisioningErrors": []
            },
            {
              "@odata.type": "#microsoft.graph.device",
              "id": "3f9fd7c3-73ad-4ce3-b053-76bb8252964d",
              "deletedDateTime": null,
              "accountEnabled": true,
              "approximateLastSignInDateTime": null,
              "complianceExpirationDateTime": null,
              "createdDateTime": "2023-11-06T06:18:26Z",
              "deviceCategory": null,
              "deviceId": "4c299165-6e8f-4b45-a5ba-c5d250a707ff",
              "deviceMetadata": null,
              "deviceOwnership": null,
              "deviceVersion": null,
              "displayName": "AdeleVence-PC",
              "domainName": null,
              "enrollmentProfileName": null,
              "enrollmentType": null,
              "externalSourceName": null,
              "isCompliant": null,
              "isManaged": null,
              "isRooted": null,
              "managementType": null,
              "manufacturer": null,
              "mdmAppId": null,
              "model": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesSyncEnabled": null,
              "operatingSystem": "windows",
              "operatingSystemVersion": "10",
              "physicalIds": [],
              "profileType": null,
              "registrationDateTime": null,
              "sourceType": null,
              "systemLabels": [],
              "trustType": null,
              "extensionAttributes": {
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
              "alternativeSecurityIds": [
                {
                  "type": 2,
                  "identityProvider": null,
                  "key": "Y3YxN2E1MWFlYw=="
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId } });

    assert(
      loggerLogSpy.calledOnceWithExactly([
        userTransformedResponse,
        groupTransformedResponse,
        deviceTransformedResponse
      ])
    );
  });

  it('should get all members of an administrative unit specified by its name when type not specified', async () => {
    sinon.stub(aadAdministrativeUnit, 'getAdministrativeUnitByDisplayName').withArgs(administrativeUnitName).resolves({ id: administrativeUnitId, displayName: administrativeUnitName });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members`) {
        return {
          value: [
            {
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
            },
            {
              "@odata.type": "#microsoft.graph.group",
              "id": "c121c70b-deb1-43f7-8298-9111bf3036b4",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2023-02-22T06:39:13Z",
              "creationOptions": [
                "Team"
              ],
              "description": "Welcome to the team that we have assembled to create the Mark 8.",
              "displayName": "Mark 8 Project Team",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "Mark8ProjectTeam@4wrvkx.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "Mark8ProjectTeam",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [],
              "renewedDateTime": "2023-02-22T06:39:13Z",
              "resourceBehaviorOptions": [
                "HideGroupInOutlook",
                "SubscribeMembersToCalendarEventsDisabled",
                "WelcomeEmailDisabled"
              ],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-3240216331-1140317873-294754434-3023450303",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": [],
              "serviceProvisioningErrors": []
            },
            {
              "@odata.type": "#microsoft.graph.device",
              "id": "3f9fd7c3-73ad-4ce3-b053-76bb8252964d",
              "deletedDateTime": null,
              "accountEnabled": true,
              "approximateLastSignInDateTime": null,
              "complianceExpirationDateTime": null,
              "createdDateTime": "2023-11-06T06:18:26Z",
              "deviceCategory": null,
              "deviceId": "4c299165-6e8f-4b45-a5ba-c5d250a707ff",
              "deviceMetadata": null,
              "deviceOwnership": null,
              "deviceVersion": null,
              "displayName": "AdeleVence-PC",
              "domainName": null,
              "enrollmentProfileName": null,
              "enrollmentType": null,
              "externalSourceName": null,
              "isCompliant": null,
              "isManaged": null,
              "isRooted": null,
              "managementType": null,
              "manufacturer": null,
              "mdmAppId": null,
              "model": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesSyncEnabled": null,
              "operatingSystem": "windows",
              "operatingSystemVersion": "10",
              "physicalIds": [],
              "profileType": null,
              "registrationDateTime": null,
              "sourceType": null,
              "systemLabels": [],
              "trustType": null,
              "extensionAttributes": {
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
              "alternativeSecurityIds": [
                {
                  "type": 2,
                  "identityProvider": null,
                  "key": "Y3YxN2E1MWFlYw=="
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitName: administrativeUnitName } });
    assert(
      loggerLogSpy.calledOnceWithExactly([
        userTransformedResponse,
        groupTransformedResponse,
        deviceTransformedResponse
      ])
    );
  });

  it('handles error when type not specified and retrieving all members of an administrative unit failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { administrativeUnitId: administrativeUnitId } }), new CommandError('An error has occurred'));
  });

  it('should get only user members of administrative unit when type is set to user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.user`) {
        return {
          value: [
            userResponseWithoutMetadata
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, type: 'user' } });

    assert(
      loggerLogSpy.calledOnceWithExactly([
        userResponseWithoutMetadata
      ])
    );
  });

  it('handles error when type is set to user and retrieving user members of an administrative unit failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.user`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { administrativeUnitId: administrativeUnitId, type: 'user' } }), new CommandError('An error has occurred'));
  });

  it('should get only group members of administrative unit when type is set to group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.group`) {
        return {
          value: [
            groupResponseWithoutMetadata
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, type: 'group' } });

    assert(
      loggerLogSpy.calledOnceWithExactly([
        groupResponseWithoutMetadata
      ])
    );
  });

  it('handles error when type is set to group and retrieving group members of an administrative unit failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.group`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { administrativeUnitId: administrativeUnitId, type: 'group' } }), new CommandError('An error has occurred'));
  });

  it('should get only device members of administrative unit when type is set to device', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.device`) {
        return {
          value: [
            deviceResponseWithoutMetadata
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, type: 'device' } });

    assert(
      loggerLogSpy.calledOnceWithExactly([
        deviceResponseWithoutMetadata
      ])
    );
  });

  it('handles error when type is set to device and retrieving device members of an administrative unit failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.device`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, { options: { administrativeUnitId: administrativeUnitId, type: 'device' } }), new CommandError('An error has occurred'));
  });

  it('should filter users of administrative unit when type is set to user and filter is specified, return only limited set of properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.user?$select=id,displayName&$expand=manager($select=displayName),drive($select=id)&$filter=givenName eq 'Pradeep'&$count=true`) {
        return {
          value: [
            limitedUserResponseWithoutMetadata
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { administrativeUnitId: administrativeUnitId, properties: "id,displayName,manager/displayName,drive/id", type: 'user', filter: "givenName eq 'Pradeep'" } });

    assert(
      loggerLogSpy.calledOnceWithExactly([
        limitedUserResponseWithoutMetadata
      ])
    );
  });
});