import commands from '../../commands';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./siteclassification-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.SITECLASSIFICATION_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
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
    auth.service = new Service('https://graph.microsoft.com');
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.patch,
      request.get,
      global.setTimeout
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
    assert.equal(command.name.startsWith(commands.SITECLASSIFICATION_SET), true);
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
        assert.equal(telemetry.name, commands.SITECLASSIFICATION_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to the Microsoft Graph', (done) => {
    auth.service = new Service('https://graph.microsoft.com');
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
    assert(find.calledWith(commands.SITECLASSIFICATION_SET));
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

  it('fails validation if the classification and defaultClassification are not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the classification is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false, defaultClassification: "HBI"
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the defaultClassification is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false, classifications: "HBI, LBI, Top Secret"
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if the required options are correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI"
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if the required options are correct and optional options are passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "https://aka.ms/pnp"
      }
    });
    assert.equal(actual, true);
  });

  it('handles Office 365 Tenant siteclassification has not been enabled', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings`) {

        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid Request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } }, (err: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("There is no previous defined site classification which can updated.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Office 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url (debug)', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings`) {

        return Promise.resolve({
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": "high,middle,low"
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

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"templateId":"a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b","values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation"}]}`) {
          updateRequestIssued = true;
      }
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, classifications: "high,middle,low", defaultClassification: "middle", usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Office 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings`) {

        return Promise.resolve({
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": "high,middle,low"
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

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"templateId":"a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b","values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation"}]}`) {
          updateRequestIssued = true;
      }
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, classifications: "high,middle,low", defaultClassification: "middle", usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Office 365 Tenant siteclassification with usage guidelines url', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings`) {

        return Promise.resolve({
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": "high,middle,low"
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

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"templateId":"a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b","values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation"}]}`) {
          updateRequestIssued = true;
      }
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { classifications: "high,middle,low", defaultClassification: "middle", usageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('updates Office 365 Tenant siteclassification with guest usage guidelines url', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings`) {

        return Promise.resolve({
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": "high,middle,low"
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

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"templateId":"a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b","values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation"}]}`) {
          updateRequestIssued = true;
      }
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, classifications: "high,middle,low", defaultClassification: "middle", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Office 365 Tenant siteclassification', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings`) {

        return Promise.resolve({
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"templateId":"a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b","values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation"}]}`) {
          updateRequestIssued = true;
      }
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { classifications: "high,middle,low", defaultClassification: "middle" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
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