import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./groupsetting-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.GROUPSETTING_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.get,
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.GROUPSETTING_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates group setting', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323`) {
        return Promise.resolve({
          "id": "c391b57d-5783-4c53-9236-cefb5c6ef323", "displayName": null, "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b", "values": [{ "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323` &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: null,
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
            {
              name: 'UsageGuidelinesUrl',
              value: 'https://contoso.sharepoint.com/sites/compliance'
            },
            {
              name: 'ClassificationList',
              value: 'HBI, MBI, LBI, GDPR'
            },
            {
              name: 'DefaultClassification',
              value: 'MBI'
            },
            {
              name: 'CustomBlockedWordsList',
              value: ''
            },
            {
              name: 'EnableMSStandardBlockedWords',
              value: 'false'
            },
            {
              name: 'ClassificationDescriptions',
              value: ''
            },
            {
              name: 'PrefixSuffixNamingRequirement',
              value: ''
            },
            {
              name: 'AllowGuestsToBeGroupOwner',
              value: 'false'
            },
            {
              name: 'AllowGuestsToAccessGroups',
              value: 'true'
            },
            {
              name: 'GuestUsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'GroupCreationAllowedGroupId',
              value: ''
            },
            {
              name: 'AllowToAddGuests',
              value: 'true'
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
        UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance',
        ClassificationList: 'HBI, MBI, LBI, GDPR',
        DefaultClassification: 'MBI'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates group setting (debug)', (done) => {
    let settingsUpdated: boolean = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323`) {
        return Promise.resolve({
          "id": "c391b57d-5783-4c53-9236-cefb5c6ef323", "displayName": null, "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b", "values": [{ "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323` &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: null,
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
            {
              name: 'UsageGuidelinesUrl',
              value: 'https://contoso.sharepoint.com/sites/compliance'
            },
            {
              name: 'ClassificationList',
              value: 'HBI, MBI, LBI, GDPR'
            },
            {
              name: 'DefaultClassification',
              value: 'MBI'
            },
            {
              name: 'CustomBlockedWordsList',
              value: ''
            },
            {
              name: 'EnableMSStandardBlockedWords',
              value: 'false'
            },
            {
              name: 'ClassificationDescriptions',
              value: ''
            },
            {
              name: 'PrefixSuffixNamingRequirement',
              value: ''
            },
            {
              name: 'AllowGuestsToBeGroupOwner',
              value: 'false'
            },
            {
              name: 'AllowGuestsToAccessGroups',
              value: 'true'
            },
            {
              name: 'GuestUsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'GroupCreationAllowedGroupId',
              value: ''
            },
            {
              name: 'AllowToAddGuests',
              value: 'true'
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        settingsUpdated = true;
        return Promise.resolve({
          displayName: null,
          id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
        UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance',
        ClassificationList: 'HBI, MBI, LBI, GDPR',
        DefaultClassification: 'MBI'
      }
    }, () => {
      try {
        assert(settingsUpdated);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('ignores global options when creating request body', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323`) {
        return Promise.resolve({
          "id": "c391b57d-5783-4c53-9236-cefb5c6ef323", "displayName": null, "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b", "values": [{ "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });
    const patchStub = sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323` &&
        JSON.stringify(opts.body) === JSON.stringify({
          displayName: null,
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
            {
              name: 'UsageGuidelinesUrl',
              value: 'https://contoso.sharepoint.com/sites/compliance'
            },
            {
              name: 'ClassificationList',
              value: 'HBI, MBI, LBI, GDPR'
            },
            {
              name: 'DefaultClassification',
              value: 'MBI'
            },
            {
              name: 'CustomBlockedWordsList',
              value: ''
            },
            {
              name: 'EnableMSStandardBlockedWords',
              value: 'false'
            },
            {
              name: 'ClassificationDescriptions',
              value: ''
            },
            {
              name: 'PrefixSuffixNamingRequirement',
              value: ''
            },
            {
              name: 'AllowGuestsToBeGroupOwner',
              value: 'false'
            },
            {
              name: 'AllowGuestsToAccessGroups',
              value: 'true'
            },
            {
              name: 'GuestUsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'GroupCreationAllowedGroupId',
              value: ''
            },
            {
              name: 'AllowToAddGuests',
              value: 'true'
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        return Promise.resolve({
          displayName: null,
          id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        verbose: true,
        output: "text",
        id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
        UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance',
        ClassificationList: 'HBI, MBI, LBI, GDPR',
        DefaultClassification: 'MBI'
      }
    }, () => {
      try {
        assert.deepEqual(patchStub.firstCall.args[0].body, {
          displayName: null,
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
            {
              name: 'UsageGuidelinesUrl',
              value: 'https://contoso.sharepoint.com/sites/compliance'
            },
            { name: 'ClassificationList', value: 'HBI, MBI, LBI, GDPR' },
            { name: 'DefaultClassification', value: 'MBI' },
            { name: 'CustomBlockedWordsList', value: '' },
            { name: 'EnableMSStandardBlockedWords', value: 'false' },
            { name: 'ClassificationDescriptions', value: '' },
            { name: 'PrefixSuffixNamingRequirement', value: '' },
            { name: 'AllowGuestsToBeGroupOwner', value: 'false' },
            { name: 'AllowGuestsToAccessGroups', value: 'true' },
            { name: 'GuestUsageGuidelinesUrl', value: '' },
            { name: 'GroupCreationAllowedGroupId', value: '' },
            { name: 'AllowToAddGuests', value: 'true' },
            { name: 'EnableGroupCreation', value: 'true' }
          ]
        });
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when no group setting with the specified id found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '62375ab9-6b52-47ed-826b-58e47e0e304c' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "request-id": "fe2491f9-53e7-407c-9a08-b92b2bf6722b",
              "date": "2018-05-11T17:06:22"
            }
          }
        }
      });
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '62375ab9-6b52-47ed-826b-58e47e0e304c' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '62375ab9-6b52-47ed-826b-58e47e0e304c' does not exist or one of its queried reference-property objects are not present.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } });
    assert.strictEqual(actual, true);
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
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