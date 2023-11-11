import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './groupsetting-set.js';

describe(commands.GROUPSETTING_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUPSETTING_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates group setting', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323`) {
        return {
          "id": "c391b57d-5783-4c53-9236-cefb5c6ef323", "displayName": null, "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b", "values": [{ "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323` &&
        JSON.stringify(opts.data) === JSON.stringify({
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

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
        UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance',
        ClassificationList: 'HBI, MBI, LBI, GDPR',
        DefaultClassification: 'MBI'
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('updates group setting (debug)', async () => {
    let settingsUpdated: boolean = false;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323`) {
        return {
          "id": "c391b57d-5783-4c53-9236-cefb5c6ef323", "displayName": null, "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b", "values": [{ "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323` &&
        JSON.stringify(opts.data) === JSON.stringify({
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
        return {
          displayName: null,
          id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
        UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance',
        ClassificationList: 'HBI, MBI, LBI, GDPR',
        DefaultClassification: 'MBI'
      }
    });
    assert(settingsUpdated);
  });

  it('ignores global options when creating request data', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323`) {
        return {
          "id": "c391b57d-5783-4c53-9236-cefb5c6ef323", "displayName": null, "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b", "values": [{ "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid request';
    });
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/c391b57d-5783-4c53-9236-cefb5c6ef323` &&
        JSON.stringify(opts.data) === JSON.stringify({
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
        return {
          displayName: null,
          id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        output: "text",
        id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
        UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance',
        ClassificationList: 'HBI, MBI, LBI, GDPR',
        DefaultClassification: 'MBI'
      }
    });
    assert.deepEqual(patchStub.firstCall.args[0].data, {
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
  });

  it('handles error when no group setting with the specified id found', async () => {
    sinon.stub(request, 'get').callsFake(() => {
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

    await assert.rejects(command.action(logger, { options: { id: '62375ab9-6b52-47ed-826b-58e47e0e304c' } } as any),
      new CommandError(`Resource '62375ab9-6b52-47ed-826b-58e47e0e304c' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });
});
