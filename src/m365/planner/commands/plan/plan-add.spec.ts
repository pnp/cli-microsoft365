import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
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
        ownerGroupName: 'My Planner Plan Group'
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
        ownerGroupName: 'My Planner Plan Group'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly adds planner plan with given title and ownerGroupId', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
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
    });
    command.action(logger, {
      options: {
        debug: false,
        title: 'My Planner Plan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    }, () => {
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

  it('correctly adds planner plan with given title and ownerGroupName', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
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
    });
    command.action(logger, {
      options: {
        debug: false,
        title: 'My Planner Plan',
        ownerGroupName: 'My Planner Plan Group'
      }
    }, () => {
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
    
    done();
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