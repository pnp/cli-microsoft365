import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./siteclassification-enable');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SITECLASSIFICATION_ENABLE, () => {
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
      request.post,
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
    assert.strictEqual(command.name.startsWith(commands.SITECLASSIFICATION_ENABLE), true);
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

  it('handles Microsoft 365 Tenant siteclassification missing DirectorySettingTemplate', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/directorySettingTemplates`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
              "displayName": "Group.Unified_not_exist",
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
                  "value": "TopSecret"
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
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
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

    cmdInstance.action({ options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\"")));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url (debug)', (done) => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/directorySettingTemplates`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
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
                  "value": "TopSecret"
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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } }, (err: any) => {
      try {
        assert(enableRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url', (done) => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/directorySettingTemplates`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
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
                  "value": "TopSecret"
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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } }, (err: any) => {
      try {
        assert(enableRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url', (done) => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/directorySettingTemplates`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
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
                  "value": "TopSecret"
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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } }, (err: any) => {
      try {
        assert(enableRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('sets Microsoft 365 Tenant siteclassification with guest usage guidelines url', (done) => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/directorySettingTemplates`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
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
                  "value": "TopSecret"
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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } }, (err: any) => {
      try {
        assert(enableRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('sets Microsoft 365 Tenant siteclassification', (done) => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/directorySettingTemplates`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
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
                  "value": "TopSecret"
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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    cmdInstance.action({ options: { debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI" } }, (err: any) => {
      try {
        assert(enableRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('Handles enabling when already enabled (conflicting errors)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/directorySettingTemplates`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
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
                  "value": "TopSecret"
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

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings` &&
        JSON.stringify(opts.body) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        return Promise.reject({
          error: {
            "error": {
              "code": "Request_BadRequest",
              "message": "A conflicting object with one or more of the specified property values is present in the directory.",
              "innerError": {
                "request-id": "fe109878-0adc-4cc8-be2a-e27a70342faa",
                "date": "2018-09-07T11:38:45"
              }
            }
          }
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: false, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI" } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`A conflicting object with one or more of the specified property values is present in the directory.`)));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });
});