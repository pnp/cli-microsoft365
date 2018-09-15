import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./groupsetting-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.GROUPSETTING_LIST, () => {
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
    assert.equal(command.name.startsWith(commands.GROUPSETTING_LIST), true);
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
        assert.equal(telemetry.name, commands.GROUPSETTING_LIST);
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

  it('lists group setting templates', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return Promise.resolve({
          "value": [
            {
              "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": ""
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": ""
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
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
        assert(cmdInstanceLogSpy.calledWith([{
          "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
          "displayName": "Group.Unified"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists group setting templates (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return Promise.resolve({
          "value": [
            {
              "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": ""
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": ""
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
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
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{
          "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
          "displayName": "Group.Unified"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('includes all properties in output type json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return Promise.resolve({
          "value": [
            {
              "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": ""
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": ""
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
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
    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
            "displayName": "Group.Unified",
            "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
            "values": [
              {
                "name": "CustomBlockedWordsList",
                "value": ""
              },
              {
                "name": "EnableMSStandardBlockedWords",
                "value": "false"
              },
              {
                "name": "ClassificationDescriptions",
                "value": ""
              },
              {
                "name": "DefaultClassification",
                "value": ""
              },
              {
                "name": "PrefixSuffixNamingRequirement",
                "value": ""
              },
              {
                "name": "AllowGuestsToBeGroupOwner",
                "value": "false"
              },
              {
                "name": "AllowGuestsToAccessGroups",
                "value": "true"
              },
              {
                "name": "GuestUsageGuidelinesUrl",
                "value": ""
              },
              {
                "name": "GroupCreationAllowedGroupId",
                "value": ""
              },
              {
                "name": "AllowToAddGuests",
                "value": "true"
              },
              {
                "name": "UsageGuidelinesUrl",
                "value": ""
              },
              {
                "name": "ClassificationList",
                "value": ""
              },
              {
                "name": "EnableGroupCreation",
                "value": "true"
              }
            ]
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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
    assert(find.calledWith(commands.GROUPSETTING_LIST));
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