import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./plan-details-get');

describe(commands.PLAN_DETAILS_GET, () => {
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
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.PLAN_DETAILS_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if neither planId nor planTitle are provided.', (done) => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both planId nor planTitle are specified', (done) => {
    const actual = command.validate({
      options: {
        planId: 'opb7bchfZUiFbVWEPL7jPGUABW7f',
        planTitle: 'MyPlan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided.', (done) => {
    const actual = command.validate({
      options: {
        planTitle: 'MyPlan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified', (done) => {
    const actual = command.validate({
      options: {
        planTitle: 'MyPlan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4',
        ownerGroupName: 'spridermvp'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        planTitle: 'MyPlan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when planId specified', (done) => {
    const actual = command.validate({
      options: {
        planId: 'opb7bchfZUiFbVWEPL7jPGUABW7f'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when planTitle and valid ownerGroupId specified', (done) => {
    const actual = command.validate({
      options: {
        planTitle: 'MyPlan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when planTitle and valid ownerGroupName specified', (done) => {
    const actual = command.validate({
      options: {
        planTitle: 'MyPlan',
        ownerGroupName: 'spridermvp'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly get planner plan details with given planId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f/details`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans('opb7bchfZUiFbVWEPL7jPGUABW7f')/details/$entity",
          "@odata.etag": "W/\"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBATCc=\"",
          "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
          "sharedWith": {
            "2ef33c97-e727-436e-ab36-3c30f06cbb21": true,
            "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e": true
          },
          "categoryDescriptions": {
            "category1": "Indoors",
            "category2": "Outdoors",
            "category3": null,
            "category4": null,
            "category5": "Needs materials",
            "category6": "Needs equipment"
          }
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      planId: 'opb7bchfZUiFbVWEPL7jPGUABW7f'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans('opb7bchfZUiFbVWEPL7jPGUABW7f')/details/$entity",
          "@odata.etag": "W/\"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBATCc=\"",
          "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
          "sharedWith": {
            "2ef33c97-e727-436e-ab36-3c30f06cbb21": true,
            "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e": true
          },
          "categoryDescriptions": {
            "category1": "Indoors",
            "category2": "Outdoors",
            "category3": null,
            "category4": null,
            "category5": "Needs materials",
            "category6": "Needs equipment"
          }
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple owner groups with same name exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null,
              "resourceProvisioningOptions": ["Team"]
            },
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null,
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        ownerGroupName: 'Team Name',
        planTitle: 'Test Plan 2'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple ownerGroups with name Team Name found: Please choose between the following IDs 00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get planner plan details with given planTitle and ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/planner/plans`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerPlan)",
          "@odata.count": 1,
          "value": [
            {
              "@odata.etag": "W/\"JzEtUZxhbiAgQEBAQEBAMEBAQEBAVEBAUCc=\"",
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

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f/details`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans('opb7bchfZUiFbVWEPL7jPGUABW7f')/details/$entity",
          "@odata.etag": "W/\"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBATCc=\"",
          "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
          "sharedWith": {
            "2ef33c97-e727-436e-ab36-3c30f06cbb21": true,
            "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e": true
          },
          "categoryDescriptions": {
            "category1": "Indoors",
            "category2": "Outdoors",
            "category3": null,
            "category4": null,
            "category5": "Needs materials",
            "category6": "Needs equipment"
          }
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      planTitle: 'My Planner Plan',
      ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans('opb7bchfZUiFbVWEPL7jPGUABW7f')/details/$entity",
          "@odata.etag": "W/\"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBATCc=\"",
          "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
          "sharedWith": {
            "2ef33c97-e727-436e-ab36-3c30f06cbb21": true,
            "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e": true
          },
          "categoryDescriptions": {
            "category1": "Indoors",
            "category2": "Outdoors",
            "category3": null,
            "category4": null,
            "category5": "Needs materials",
            "category6": "Needs equipment"
          }
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get planner plan details with given plantitle and ownerGroupName', (done) => {
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
              "@odata.etag": "W/\"JzEtUZxhbiAgQEBAQEBAMEBAQEBAVEBAUCc=\"",
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

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f/details`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans('opb7bchfZUiFbVWEPL7jPGUABW7f')/details/$entity",
          "@odata.etag": "W/\"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBATCc=\"",
          "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
          "sharedWith": {
            "2ef33c97-e727-436e-ab36-3c30f06cbb21": true,
            "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e": true
          },
          "categoryDescriptions": {
            "category1": "Indoors",
            "category2": "Outdoors",
            "category3": null,
            "category4": null,
            "category5": "Needs materials",
            "category6": "Needs equipment"
          }
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });


    const options: any = {
      debug: false,
      planTitle: 'My Planner Plan',
      ownerGroupName: 'spridermvp'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans('opb7bchfZUiFbVWEPL7jPGUABW7f')/details/$entity",
          "@odata.etag": "W/\"JzEtUGxhbkRldGFpbHMgQEBAQEBAQEBAQEBAQEBATCc=\"",
          "id": "opb7bchfZUiFbVWEPL7jPGUABW7f",
          "sharedWith": {
            "2ef33c97-e727-436e-ab36-3c30f06cbb21": true,
            "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e": true
          },
          "categoryDescriptions": {
            "category1": "Indoors",
            "category2": "Outdoors",
            "category3": null,
            "category4": null,
            "category5": "Needs materials",
            "category6": "Needs equipment"
          }
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('failed validation when multiple planner plans are found with same name and ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/planner/plans`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerPlan)",
          "@odata.count": 1,
          "value": [
            {
              "@odata.etag": "W/\"JzEtUZxhbiAgQEBAQEBAMEBAQEBAVEBAUCc=\"",
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
            },
            {
              "@odata.etag": "W/\"JzEtUZxhbiAgQEBAQEBAMEBAQEBAVEBAUCc=\"",
              "createdDateTime": "2021-03-10T17:39:43.1045549Z",
              "owner": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
              "title": "My Planner Plan",
              "id": "KEBXwYWi8025K93fSZKwOZgAGULL",
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

    command.action(logger, {
      options: {
        debug: true,
        planTitle: 'My Planner Plan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple plans with name My Planner Plan found: opb7bchfZUiFbVWEPL7jPGUABW7f,KEBXwYWi8025K93fSZKwOZgAGULL`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when ownerGroupName not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        planTitle: 'My Planner Plan',
        ownerGroupName: 'foo'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified ownerGroup does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get owmnerGroup when ownerGroup does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified ownerGroup does not exist');
    });

    command.action(logger, {
      options: {
        debug: true,
        planTitle: 'My Planner Plan',
        ownerGroupName: 'Team Name'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified ownerGroup does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('correctly handles no plan found with given ownerGroupId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
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
      title: 'My Planner Plan',
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