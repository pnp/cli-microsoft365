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
import command, { options } from './plan-list.js';

describe(commands.PLAN_LIST, () => {
  const ownerGroupId = '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4';
  const ownerGroupName = 'spridermvp';
  const rosterId = 'FeMZFDoK8k2oWmuGE-XFHZcAEwtn';
  const groupsResponse = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
    "value": [
      {
        "id": ownerGroupId,
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2021-01-23T17:58:03Z",
        "creationOptions": [
          "Team",
          "ExchangeProvisioningFlags:3552"
        ],
        "description": "Check here for organization announcements and important info.",
        "displayName": "spridermvp",
        "expirationDateTime": null,
        "groupTypes": [
          "Unified"
        ],
        "isAssignableToRole": null,
        "mail": "spridermvp@spridermvp.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "spridermvp",
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
        "proxyAddresses": [
          "SPO:SPO_fe66856a-ca60-457c-9215-cef02b57bf01@SPO_b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923",
          "SMTP:spridermvp@spridermvp.onmicrosoft.com"
        ],
        "renewedDateTime": "2021-01-23T17:58:03Z",
        "resourceBehaviorOptions": [
          "HideGroupInOutlook",
          "SubscribeMembersToCalendarEventsDisabled",
          "WelcomeEmailDisabled"
        ],
        "resourceProvisioningOptions": [
          "Team"
        ],
        "securityEnabled": false,
        "securityIdentifier": "S-1-12-1-591283152-1211030634-3876408987-3035217063",
        "theme": null,
        "visibility": "Public",
        "onPremisesProvisioningErrors": []
      }
    ]
  };
  const planResponse = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerPlan)",
    "@odata.count": 1,
    "value": [
      {
        "createdDateTime": "2021-03-10T17:39:43.1045549Z",
        "owner": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
        "title": "My Planner Plan",
        "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "eded3a2a-8f01-40aa-998a-e4f02ec693ba"
          },
          "application": {
            "displayName": null,
            "id": "31359c7f-bd7e-475c-86db-fdb8c937548e"
          }
        }
      }
    ]
  };
  const formattedResponse = [{
    "createdDateTime": "2021-03-10T17:39:43.1045549Z",
    "owner": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
    "title": "My Planner Plan",
    "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "eded3a2a-8f01-40aa-998a-e4f02ec693ba"
      },
      "application": {
        "displayName": null,
        "id": "31359c7f-bd7e-475c-86db-fdb8c937548e"
      }
    }
  }];

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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PLAN_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner']);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', () => {
    const actual = commandOptionsSchema.safeParse({
      ownerGroupId: 'invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName nor rosterId are provided.', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when ownerGroupId, rosterId and ownerGroupName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      ownerGroupId: ownerGroupId,
      ownerGroupName: ownerGroupName,
      rosterId: rosterId
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when valid ownerGroupId specified', () => {
    const actual = commandOptionsSchema.safeParse({
      ownerGroupId: ownerGroupId
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when valid ownerGroupName specified', () => {
    const actual = commandOptionsSchema.safeParse({
      ownerGroupName: ownerGroupName
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when valid rosterId specified', () => {
    const actual = commandOptionsSchema.safeParse({
      rosterId: rosterId
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      ownerGroupId: ownerGroupId,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('correctly list planner plans with given ownerGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return planResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ ownerGroupId: ownerGroupId }) });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('correctly list planner plans with given ownerGroupName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${ownerGroupName}'&$select=id`) {
        return groupsResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return planResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ ownerGroupName: ownerGroupName }) });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('correctly list planner plans with given roster', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${rosterId}/plans`) {
        return planResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ rosterId: rosterId }) });
    assert(loggerLogSpy.calledWith(formattedResponse));
  });

  it('correctly handles no plan found with given ownerGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerPlan)",
          "@odata.count": 0,
          "value": []
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ ownerGroupId: ownerGroupId }) });
    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred.'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ ownerGroupId: ownerGroupId }) }), new CommandError("An error has occurred."));
  });
});
