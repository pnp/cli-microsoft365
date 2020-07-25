import commands from '../../commands';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./siteclassification-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SITECLASSIFICATION_SET, () => {
  let log: string[];
  let cmdInstance: any;

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
  });

  afterEach(() => {
    Utils.restore([
      request.patch,
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
    assert.strictEqual(command.name.startsWith(commands.SITECLASSIFICATION_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('fails validation if none of the options are specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if at least one option is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false, classifications: "Confidential"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if all options are passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "https://aka.ms/pnp", guestUsageGuidelinesUrl: "https://aka.ms/pnp"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('handles Microsoft 365 Tenant siteclassification has not been enabled', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("There is no previous defined site classification which can updated.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant usage guidelines url and guest usage guidelines url (debug)', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: true, usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant usage guidelines url and guest usage guidelines url', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: false, usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant usage guidelines url', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { usageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant guest usage guidelines url', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: false, guestUsageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant siteclassification', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"top secret,high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { classifications: "top secret,high,middle,low" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant default classification', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"low"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;
        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { defaultClassification: "low" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant siteclassification and default classification', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"high"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"area 51,high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { classifications: "area 51,high,middle,low", defaultClassification: "high" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates Microsoft 365 Tenant siteclassification, default classification, usage guidelines url and guest usage guidelines url', (done) => {
    let updateRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
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
      if (opts.url === `https://graph.microsoft.com/beta/settings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"high"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"area 51,high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { classifications: "area 51,high,middle,low", defaultClassification: "high", usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } }, (err: any) => {
      try {
        assert(updateRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});