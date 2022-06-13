import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./plan-list');

describe(commands.PLAN_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PLAN_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner']);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided.', (done) => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified', (done) => {
    const actual = command.validate({
      options: {
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4',
        ownerGroupName: 'spridermvp'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid ownerGroupId specified', (done) => {
    const actual = command.validate({
      options: {
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid ownerGroupName specified', (done) => {
    const actual = command.validate({
      options: {
        ownerGroupName: 'spridermvp'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly list planner plans with given ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=ID') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
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
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/planner/plans`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith([{
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
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly list planner plans with given ownerGroupName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
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
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/planner/plans`) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      ownerGroupName: 'spridermvp'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith([{
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
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when using app only access token', (done) => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    command.action(logger, {
      options: {
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('This command does not support application permissions.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no plan found with given ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=ID') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
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
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/planner/plans`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerPlan)",
          "@odata.count": 0,
          "value": []
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject("An error has occurred.");
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});