import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./plan-add');

describe(commands.PLANNER_PLAN_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
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
    Utils.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PLANNER_PLAN_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        title: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided.', (done) => {
    const actual = command.validate({
      options: {
        title: 'My Planner Plan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified', (done) => {
    const actual = command.validate({
      options: {
        title: 'My Planner Plan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4',
        ownerGroupName: 'spridermvp'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid title and ownerGroupId specified', (done) => {
    const actual = command.validate({
      options: {
        title: 'My Planner Plan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid title and ownerGroupName specified', (done) => {
    const actual = command.validate({
      options: {
        title: 'My Planner Plan',
        ownerGroupName: 'spridermvp'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly adds planner plan with given title with available ownerGroupId', (done) => {
    let getFakeOneGroupFound = (opts: any) => {
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
      return Promise.reject('Invalid request');
    }

    let postFakeAddPlannerPlan = (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
    }

    sinon.stub(request, 'get').callsFake(getFakeOneGroupFound);
    sinon.stub(request, 'post').callsFake(postFakeAddPlannerPlan);

    let options: any = {
      debug: false,
      title: 'My Planner Plan',
      ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
    }

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly adds planner plan with given title with available ownerGroupName', (done) => {
    let getFakeOneGroupFound = (opts: any) => {
      if ((opts.url as string).indexOf('/groups?$filter=DisplayName') > -1) {
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
      return Promise.reject('Invalid request');
    }

    let postFakeAddPlannerPlan = (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
    }

    sinon.stub(request, 'get').callsFake(getFakeOneGroupFound);
    sinon.stub(request, 'post').callsFake(postFakeAddPlannerPlan);

    let options: any = {
      debug: false,
      title: 'My Planner Plan',
      ownerGroupName: 'spridermvp'
    }

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when ownerGroupName not found', (done) => {
    let getFakeNoGroupsFound = (opts: any) => {
      if ((opts.url as string).indexOf('/groups?$filter=DisplayName') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": []
        });
      }
      return Promise.reject('Invalid request');
    }

    sinon.stub(request, 'get').callsFake(getFakeNoGroupsFound);

    let options: any = {
      debug: false,
      title: 'My Planner Plan',
      ownerGroupName: 'foo'
    }

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(`Owner group not found with name foo`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when ownerGroupId not found', (done) => {
    let getFakeNoGroupsFound = (opts: any) => {
      if ((opts.url as string).indexOf('/groups?$filter=ID') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": []
        });
      }
      return Promise.reject('Invalid request');
    }

    sinon.stub(request, 'get').callsFake(getFakeNoGroupsFound);

    let options: any = {
      debug: false,
      title: 'My Planner Plan',
      ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce123'
    }

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(`Owner group not found with id 233e43d0-dc6a-482e-9b4e-0de7a7bce123`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when more than one ownerGroupName found', (done) => {

    let getFakeMoreThanOneGroup = (opts: any) => {
      if ((opts.url as string).indexOf('/groups?$filter=DisplayName') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
              {
                  "id": "134f5bf2-238b-4b89-b855-cba02b1ffc60",
                  "deletedDateTime": null,
                  "classification": null,
                  "createdDateTime": "2020-09-22T21:41:28Z",
                  "creationOptions": [
                      "ProvisionGroupHomepage",
                      "HubSiteId:00000000-0000-0000-0000-000000000000",
                      "SPSiteLanguage:1033"
                  ],
                  "description": "dev",
                  "displayName": "dev",
                  "expirationDateTime": null,
                  "groupTypes": [
                      "Unified"
                  ],
                  "isAssignableToRole": null,
                  "mail": "dev@spridermvp.onmicrosoft.com",
                  "mailEnabled": true,
                  "mailNickname": "dev",
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
                      "SPO:SPO_8faf2144-5902-4bcc-b91f-a5d3b2c8d7ea@SPO_b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923",
                      "SMTP:dev@spridermvp.onmicrosoft.com"
                  ],
                  "renewedDateTime": "2020-09-22T21:41:28Z",
                  "resourceBehaviorOptions": [],
                  "resourceProvisioningOptions": [],
                  "securityEnabled": false,
                  "securityIdentifier": "S-1-12-1-323967986-1267278731-2697680312-1627135787",
                  "theme": null,
                  "visibility": "Private",
                  "onPremisesProvisioningErrors": []
              },
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
      return Promise.reject('Invalid request');
    }

    sinon.stub(request, 'get').callsFake(getFakeMoreThanOneGroup);

    let options: any = {
      debug: false,
      title: 'My Planner Plan',
      ownerGroupName: 'spridermvp'
    }

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(`More than one groups found with name spridermvp`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when more than one ownerGroupId found', (done) => {

    let getFakeMoreThanOneGroup = (opts: any) => {
      if ((opts.url as string).indexOf('/groups?$filter=ID') > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
              {
                  "id": "134f5bf2-238b-4b89-b855-cba02b1ffc60",
                  "deletedDateTime": null,
                  "classification": null,
                  "createdDateTime": "2020-09-22T21:41:28Z",
                  "creationOptions": [
                      "ProvisionGroupHomepage",
                      "HubSiteId:00000000-0000-0000-0000-000000000000",
                      "SPSiteLanguage:1033"
                  ],
                  "description": "dev",
                  "displayName": "dev",
                  "expirationDateTime": null,
                  "groupTypes": [
                      "Unified"
                  ],
                  "isAssignableToRole": null,
                  "mail": "dev@spridermvp.onmicrosoft.com",
                  "mailEnabled": true,
                  "mailNickname": "dev",
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
                      "SPO:SPO_8faf2144-5902-4bcc-b91f-a5d3b2c8d7ea@SPO_b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923",
                      "SMTP:dev@spridermvp.onmicrosoft.com"
                  ],
                  "renewedDateTime": "2020-09-22T21:41:28Z",
                  "resourceBehaviorOptions": [],
                  "resourceProvisioningOptions": [],
                  "securityEnabled": false,
                  "securityIdentifier": "S-1-12-1-323967986-1267278731-2697680312-1627135787",
                  "theme": null,
                  "visibility": "Private",
                  "onPremisesProvisioningErrors": []
              },
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
      return Promise.reject('Invalid request');
    }

    sinon.stub(request, 'get').callsFake(getFakeMoreThanOneGroup);

    let options: any = {
      debug: false,
      title: 'My Planner Plan',
      ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce123'
    }

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(`More than one groups found with id 233e43d0-dc6a-482e-9b4e-0de7a7bce123`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
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