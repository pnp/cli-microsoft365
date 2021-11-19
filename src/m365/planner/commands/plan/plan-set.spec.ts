import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import * as planGetCommand from '../plan/plan-get';
const command: Command = require('./plan-set');

describe(commands.PLAN_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.patch,
      Cli.executeCommandWithOutput
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
    assert.strictEqual(command.name.startsWith(commands.PLAN_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });


  it('fails validation if newTitle is not provided.', (done) => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither id nor title are provided.', (done) => {
    const actual = command.validate({
      options: {
        newTitle: 'MyNewPlan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both id and title are specified', (done) => {
    const actual = command.validate({
      options: {
        newTitle: 'MyNewPlan',
        id: 'opb7bchfZUiFbVWEPL7jPGUABW7f',
        title: 'MyPlan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided when title is set.', (done) => {
    const actual = command.validate({
      options: {
        newTitle: 'MyNewPlan',
        title: 'MyPlan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified when title is set', (done) => {
    const actual = command.validate({
      options: {
        newTitle: 'MyNewPlan',
        title: 'MyPlan',
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
        newTitle: 'MyNewPlan',
        title: 'MyPlan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when id specified', (done) => {
    const actual = command.validate({
      options: {
        newTitle: 'MyNewPlan',
        id: 'opb7bchfZUiFbVWEPL7jPGUABW7f'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when title and valid ownerGroupId specified', (done) => {
    const actual = command.validate({
      options: {
        newTitle: 'MyNewPlan',
        title: 'MyPlan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when title and valid ownerGroupName specified', (done) => {
    const actual = command.validate({
      options: {
        newTitle: 'MyNewPlan',
        title: 'MyPlan',
        ownerGroupName: 'spridermvp'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });


  it('correctly set new title when plan id is specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans/$entity",
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
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f' &&
        opts.data &&
        opts.data.title !== undefined) {
        if (opts.data.title === 'MyNewPlan') {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug : false,
        id: 'opb7bchfZUiFbVWEPL7jPGUABW7f',
        newTitle: 'MyNewPlan'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly set new title when plan id is specified (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans/$entity",
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
        });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f' &&
        opts.data &&
        opts.data.title !== undefined) {
        if (opts.data.title === 'MyNewPlan') {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug : true,
        id: 'opb7bchfZUiFbVWEPL7jPGUABW7f',
        newTitle: 'MyNewPlan'
      }
    }, (err?: any) => {
      try {
        assert(loggerLogToStderrSpy.called);
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly set new title when plan title and ownerGroupId is specified', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === planGetCommand) {
        return Promise.resolve({
          stdout: '{"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#planner/plans/$entity)","@odata.etag":"W/\\"JzEtUZxhbiAgQEBAQEBAMEBAQEBAVEBAUCc=\\"","createdDateTime":"2021-03-10T17:39:43.1045549Z","owner":"233e43d0-dc6a-482e-9b4e-0de7a7bce9b4","title":"My Planner Plan","id":"opb7bchfZUiFbVWEPL7jPGUABW7f","createdBy":{"user":{"displayName":null,"id":"eded3a2a-8f01-40aa-998a-e4f02ec693ba"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}}'
        });
      }

      return Promise.reject(`Invalid request`);
    });

    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f' &&
        opts.data &&
        opts.data.title !== undefined) {
        if (opts.data.title === 'MyNewPlan') {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug : false,
        title: 'MyPlan',
        ownerGroupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4',
        newTitle: 'MyNewPlan'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly set new title when plan title and ownerGroupName is specified', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === planGetCommand) {
        return Promise.resolve({
          stdout: '{"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#planner/plans/$entity)","@odata.etag":"W/\\"JzEtUZxhbiAgQEBAQEBAMEBAQEBAVEBAUCc=\\"","createdDateTime":"2021-03-10T17:39:43.1045549Z","owner":"233e43d0-dc6a-482e-9b4e-0de7a7bce9b4","title":"My Planner Plan","id":"opb7bchfZUiFbVWEPL7jPGUABW7f","createdBy":{"user":{"displayName":null,"id":"eded3a2a-8f01-40aa-998a-e4f02ec693ba"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}}'
        });
      }

      return Promise.reject(`Invalid request`);
    });

    sinon.stub(request, 'patch').callsFake(opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/planner/plans/opb7bchfZUiFbVWEPL7jPGUABW7f' &&
        opts.data &&
        opts.data.title !== undefined) {
        if (opts.data.title === 'MyNewPlan') {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    command.action(logger, {
      options: {
        debug : false,
        title: 'MyPlan',
        ownerGroupName: 'MyGroup',
        newTitle: 'MyNewPlan'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
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