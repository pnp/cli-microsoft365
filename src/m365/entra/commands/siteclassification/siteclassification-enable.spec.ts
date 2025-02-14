import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './siteclassification-enable.js';

describe(commands.SITECLASSIFICATION_ENABLE, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITECLASSIFICATION_ENABLE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles Microsoft 365 Tenant siteclassification missing DirectorySettingTemplate', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return {
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
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } } as any),
      new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\""));
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url (debug)', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return {
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
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url and guest usage guidelines url', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return {
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
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification with usage guidelines url', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return {
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
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification with guest usage guidelines url', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return {
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
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/sppnp"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", guestUsageGuidelinesUrl: "http://aka.ms/sppnp" } } as any);
    assert(enableRequestIssued);
  });

  it('sets Microsoft 365 Tenant siteclassification', async () => {
    let enableRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return {
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
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        enableRequestIssued = true;

        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI" } } as any);
    assert(enableRequestIssued);
  });

  it('Handles enabling when already enabled (conflicting errors)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return {
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
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` && JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList"},{"name":"EnableMSStandardBlockedWords"},{"name":"ClassificationDescriptions"},{"name":"DefaultClassification","value":"HBI"},{"name":"PrefixSuffixNamingRequirement"},{"name":"AllowGuestsToBeGroupOwner"},{"name":"AllowGuestsToAccessGroups"},{"name":"GuestUsageGuidelinesUrl"},{"name":"GroupCreationAllowedGroupId"},{"name":"AllowToAddGuests"},{"name":"UsageGuidelinesUrl"},{"name":"ClassificationList","value":"HBI, LBI, Top Secret"},{"name":"EnableGroupCreation"}],"templateId":"d20c475c-6f96-449a-aee8-08146be187d3"}`) {
        throw {
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
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, { options: { classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI" } } as any),
      new CommandError(`A conflicting object with one or more of the specified property values is present in the directory.`));
  });
});
