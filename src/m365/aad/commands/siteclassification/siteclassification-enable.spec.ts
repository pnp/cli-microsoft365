import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./siteclassification-enable');

describe(commands.SITECLASSIFICATION_ENABLE, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITECLASSIFICATION_ENABLE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles Microsoft 365 Tenant siteclassification missing DirectorySettingTemplate', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
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

    await assert.rejects(command.action(logger, { options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } } as any),
      new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\""));
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url (debug)', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    await command.action(logger, { options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification with guest usage guidelines url', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return Promise.resolve();
      }

      return Promise.reject();
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI" } } as any);
    assert(enableRequestIssued);
  });

  it('Handles enabling when already enabled (conflicting errors)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
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

    await assert.rejects(command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI" } } as any),
      new CommandError(`A conflicting object with one or more of the specified property values is present in the directory.`));
  });
});
