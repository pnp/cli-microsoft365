import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-app-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_APP_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service();
    telemetry = null;
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TEAMS_APP_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.TEAMS_APP_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to Microsoft Graph', (done) => {
    auth.service = new Service();
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to the Microsoft Graph first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
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

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid'
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if the teamId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation when the teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55'
      }
    });
    assert.equal(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TEAMS_APP_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});