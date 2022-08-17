import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./groupsetting-set');

describe(commands.GROUPSETTING_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinonUtil.restore([
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

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
        UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance',
        ClassificationList: 'HBI, MBI, LBI, GDPR',
        DefaultClassification: 'MBI'
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);
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
        return Promise.resolve({
          displayName: null,
          id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
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

  it('ignores global options when creating request data', (done) => {
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
        return Promise.resolve({
          displayName: null,
          id: 'c391b57d-5783-4c53-9236-cefb5c6ef323',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when no group setting with the specified id found', (done) => {
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

    command.action(logger, { options: { debug: false, id: '62375ab9-6b52-47ed-826b-58e47e0e304c' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '62375ab9-6b52-47ed-826b-58e47e0e304c' does not exist or one of its queried reference-property objects are not present.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});