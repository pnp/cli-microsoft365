import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./groupsetting-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.GROUPSETTING_ADD, () => {
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
    assert.strictEqual(command.name.startsWith(commands.GROUPSETTING_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds group setting using default template setting values', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return Promise.resolve({
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.body) === JSON.stringify({
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
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
              name: 'DefaultClassification',
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
              name: 'UsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'ClassificationList',
              value: ''
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        return Promise.resolve({
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds group setting using default template setting values (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return Promise.resolve({
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.body) === JSON.stringify({
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
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
              name: 'DefaultClassification',
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
              name: 'UsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'ClassificationList',
              value: ''
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        return Promise.resolve({
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds group setting using the specified values', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return Promise.resolve({
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.body) === JSON.stringify({
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
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b', UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance', ClassificationList: 'HBI, MBI, LBI, GDPR', DefaultClassification: 'MBI' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('ignores global options when creating request body', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return Promise.resolve({
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        });
      }

      return Promise.reject('Invalid request');
    });
    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.body) === JSON.stringify({
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
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, verbose: true, output: "text", templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b', UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance', ClassificationList: 'HBI, MBI, LBI, GDPR', DefaultClassification: 'MBI' } }, () => {
      try {
        assert.deepEqual(postStub.firstCall.args[0].body, {
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

  it('handles error when no template with the specified id found', (done) => {
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

  it('handles error when group setting already exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return Promise.resolve({
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        error: {
          "error": {
            "code": "Request_BadRequest",
            "message": "A conflicting object with one or more of the specified property values is present in the directory.",
            "innerError": {
              "request-id": "7b7eacbb-3b0e-4758-be20-6410957e42d6",
              "date": "2018-05-11T17:10:34"
            }
          }
        }
      });
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`A conflicting object with one or more of the specified property values is present in the directory.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the templateId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { templateId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the templateId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { templateId: '68be84bf-a585-4776-80b3-30aa5207aa22' } });
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