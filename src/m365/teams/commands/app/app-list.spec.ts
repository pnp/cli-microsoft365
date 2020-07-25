import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./app-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_APP_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_APP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists Microsoft Teams apps in the organization app catalog', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=distributionMethod eq 'organization'`) {
        return Promise.resolve({
          "value": [
            {
              "id": "7131a36d-bb5f-46b8-bb40-0b199a3fad74",
              "externalId": "4f0cd7c8-995e-4868-812d-d1d402a81eca",
              "displayName": "WsInfo",
              "distributionMethod": "organization"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "7131a36d-bb5f-46b8-bb40-0b199a3fad74",
            "displayName": "WsInfo",
            "distributionMethod": "organization"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft Teams apps in the organization app catalog and Microsoft Teams store', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return Promise.resolve({
          "value": [
            {
              "id": "012be6ac-6f34-4ffa-9344-b857f7bc74e1",
              "externalId": null,
              "displayName": "Pickit Images",
              "distributionMethod": "store"
            },
            {
              "id": "01b22ab6-c657-491c-97a0-d745bea11269",
              "externalId": null,
              "displayName": "Hootsuite",
              "distributionMethod": "store"
            },
            {
              "id": "02d14659-a28b-4007-8544-b279c0d3628b",
              "externalId": null,
              "displayName": "Pivotal Tracker",
              "distributionMethod": "store"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { all: true, debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "012be6ac-6f34-4ffa-9344-b857f7bc74e1",
            "displayName": "Pickit Images",
            "distributionMethod": "store"
          },
          {
            "id": "01b22ab6-c657-491c-97a0-d745bea11269",
            "displayName": "Hootsuite",
            "distributionMethod": "store"
          },
          {
            "id": "02d14659-a28b-4007-8544-b279c0d3628b",
            "displayName": "Pivotal Tracker",
            "distributionMethod": "store"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists organization\'s apps installed in a team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6f6fd3f7-9ba5-4488-bbe6-a789004d0d55/installedApps?$expand=teamsApp&$filter=teamsApp/distributionMethod eq 'organization'`) {
        return Promise.resolve({
          "value": [{
            "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyNiOGNjZjNmNC04NGVlLTRlNjItODJkMC1iZjZiZjk1YmRiODM=", "teamsApp": { "id": "b8ccf3f4-84ee-4e62-82d0-bf6bf95bdb83", "externalId": "912e9d76-1794-414f-82fd-e5b60fab731b", "displayName": "HelloWorld", "distributionMethod": "organization" }
          }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyNiOGNjZjNmNC04NGVlLTRlNjItODJkMC1iZjZiZjk1YmRiODM=",
            "displayName": "HelloWorld",
            "distributionMethod": "organization"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all apps installed in a team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6f6fd3f7-9ba5-4488-bbe6-a789004d0d55/installedApps?$expand=teamsApp`) {
        return Promise.resolve({
          "value": [
            {
              "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyNiOGNjZjNmNC04NGVlLTRlNjItODJkMC1iZjZiZjk1YmRiODM=", "teamsApp": { "id": "b8ccf3f4-84ee-4e62-82d0-bf6bf95bdb83", "externalId": "912e9d76-1794-414f-82fd-e5b60fab731b", "displayName": "HelloWorld", "distributionMethod": "organization" }
            },
            {
              "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyMwZDgyMGVjZC1kZWYyLTQyOTctYWRhZC03ODA1NmNkZTdjNzg=", "teamsApp": { "id": "0d820ecd-def2-4297-adad-78056cde7c78", "externalId": null, "displayName": "OneNote", "distributionMethod": "store" }
            },
            {
              "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=", "teamsApp": { "id": "14d6962d-6eeb-4f48-8890-de55454bb136", "externalId": null, "displayName": "Activity", "distributionMethod": "store" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55', all: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyNiOGNjZjNmNC04NGVlLTRlNjItODJkMC1iZjZiZjk1YmRiODM=",
            "displayName": "HelloWorld",
            "distributionMethod": "organization"
          },
          {
            "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyMwZDgyMGVjZC1kZWYyLTQyOTctYWRhZC03ODA1NmNkZTdjNzg=",
            "displayName": "OneNote",
            "distributionMethod": "store"
          },
          {
            "id": "NmY2ZmQzZjctOWJhNS00NDg4LWJiZTYtYTc4OTAwNGQwZDU1IyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=",
            "displayName": "Activity",
            "distributionMethod": "store"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all properties for output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=distributionMethod eq 'organization'`) {
        return Promise.resolve({
          "value": [
            {
              "id": "7131a36d-bb5f-46b8-bb40-0b199a3fad74",
              "externalId": "4f0cd7c8-995e-4868-812d-d1d402a81eca",
              "displayName": "WsInfo",
              "distributionMethod": "organization"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "7131a36d-bb5f-46b8-bb40-0b199a3fad74",
            "externalId": "4f0cd7c8-995e-4868-812d-d1d402a81eca",
            "displayName": "WsInfo",
            "distributionMethod": "organization"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving apps', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});