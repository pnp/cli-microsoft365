import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './groupsettingtemplate-get.js';
import { settingsNames } from '../../../../settingsNames.js';
import aadCommands from '../../aadCommands.js';

describe(commands.GROUPSETTINGTEMPLATE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = cli.getCommandInfo(command);
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
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUPSETTINGTEMPLATE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.GROUPSETTINGTEMPLATE_GET]);
  });

  it('retrieves group setting template by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return { "value": [{ "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }] }, { "id": "08d542b9-071f-4e16-94b0-74abb372e3d9", "deletedDateTime": null, "displayName": "Group.Unified.Guest", "description": "Settings for a specific Unified Group", "values": [{ "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in a specific Unified Group." }] }, { "id": "4bc7f740-180e-4586-adb6-38b2e9024e6b", "deletedDateTime": null, "displayName": "Application", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide application behavior.\n      ", "values": [{ "name": "EnableAccessCheckForPrivilegedApplicationUpdates", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if access check for application privileged updates is turned on." }] }, { "id": "898f1161-d651-43d1-805c-3b0b388a9fc2", "deletedDateTime": null, "displayName": "Custom Policy Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide custom policy settings.\n      ", "values": [{ "name": "CustomConditionalAccessPolicyUrl", "type": "System.String", "defaultValue": "", "description": "Custom conditional access policy url." }] }, { "id": "5cf42378-d67d-4f36-ba46-e8b86229381d", "deletedDateTime": null, "displayName": "Password Rule Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide password rule settings.\n      ", "values": [{ "name": "LockoutDurationInSeconds", "type": "System.Int32", "defaultValue": "60", "description": "The duration in seconds of the initial lockout period." }, { "name": "LockoutThreshold", "type": "System.Int32", "defaultValue": "10", "description": "The number of failed login attempts before the first lockout period begins." }, { "name": "BannedPasswordList", "type": "System.String", "defaultValue": "", "description": "A tab-delimited banned password list." }, { "name": "EnableBannedPasswordCheck", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if the banned password check is turned on." }] }, { "id": "80661d51-be2f-4d46-9713-98a2fcaec5bc", "deletedDateTime": null, "displayName": "Prohibited Names Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names settings.\n      ", "values": [{ "name": "CustomBlockedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to block for application display names." }, { "name": "CustomBlockedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of reserved words to block for application display names." }] }, { "id": "aad3907d-1d1a-448b-b3ef-7bf7f63db63b", "deletedDateTime": null, "displayName": "Prohibited Names Restricted Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names restricted settings.\n      ", "values": [{ "name": "CustomAllowedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to allow for application display names." }, { "name": "CustomAllowedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of whole reserved words to allow for application display names." }, { "name": "DoNotValidateAgainstTrademark", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if prohibited names validation against trademark global list is disabled." }] }] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '62375ab9-6b52-47ed-826b-58e47e0e304b' } });
    assert(loggerLogSpy.calledWith({ "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }] }));
  });

  it('retrieves group setting template by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return { "value": [{ "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }] }, { "id": "08d542b9-071f-4e16-94b0-74abb372e3d9", "deletedDateTime": null, "displayName": "Group.Unified.Guest", "description": "Settings for a specific Unified Group", "values": [{ "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in a specific Unified Group." }] }, { "id": "4bc7f740-180e-4586-adb6-38b2e9024e6b", "deletedDateTime": null, "displayName": "Application", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide application behavior.\n      ", "values": [{ "name": "EnableAccessCheckForPrivilegedApplicationUpdates", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if access check for application privileged updates is turned on." }] }, { "id": "898f1161-d651-43d1-805c-3b0b388a9fc2", "deletedDateTime": null, "displayName": "Custom Policy Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide custom policy settings.\n      ", "values": [{ "name": "CustomConditionalAccessPolicyUrl", "type": "System.String", "defaultValue": "", "description": "Custom conditional access policy url." }] }, { "id": "5cf42378-d67d-4f36-ba46-e8b86229381d", "deletedDateTime": null, "displayName": "Password Rule Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide password rule settings.\n      ", "values": [{ "name": "LockoutDurationInSeconds", "type": "System.Int32", "defaultValue": "60", "description": "The duration in seconds of the initial lockout period." }, { "name": "LockoutThreshold", "type": "System.Int32", "defaultValue": "10", "description": "The number of failed login attempts before the first lockout period begins." }, { "name": "BannedPasswordList", "type": "System.String", "defaultValue": "", "description": "A tab-delimited banned password list." }, { "name": "EnableBannedPasswordCheck", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if the banned password check is turned on." }] }, { "id": "80661d51-be2f-4d46-9713-98a2fcaec5bc", "deletedDateTime": null, "displayName": "Prohibited Names Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names settings.\n      ", "values": [{ "name": "CustomBlockedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to block for application display names." }, { "name": "CustomBlockedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of reserved words to block for application display names." }] }, { "id": "aad3907d-1d1a-448b-b3ef-7bf7f63db63b", "deletedDateTime": null, "displayName": "Prohibited Names Restricted Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names restricted settings.\n      ", "values": [{ "name": "CustomAllowedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to allow for application display names." }, { "name": "CustomAllowedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of whole reserved words to allow for application display names." }, { "name": "DoNotValidateAgainstTrademark", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if prohibited names validation against trademark global list is disabled." }] }] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, displayName: 'Group.Unified' } });
    assert(loggerLogSpy.calledWith({ "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }] }));
  });

  it('returns error when no template with the specified id found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return { "value": [{ "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }] }, { "id": "08d542b9-071f-4e16-94b0-74abb372e3d9", "deletedDateTime": null, "displayName": "Group.Unified.Guest", "description": "Settings for a specific Unified Group", "values": [{ "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in a specific Unified Group." }] }, { "id": "4bc7f740-180e-4586-adb6-38b2e9024e6b", "deletedDateTime": null, "displayName": "Application", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide application behavior.\n      ", "values": [{ "name": "EnableAccessCheckForPrivilegedApplicationUpdates", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if access check for application privileged updates is turned on." }] }, { "id": "898f1161-d651-43d1-805c-3b0b388a9fc2", "deletedDateTime": null, "displayName": "Custom Policy Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide custom policy settings.\n      ", "values": [{ "name": "CustomConditionalAccessPolicyUrl", "type": "System.String", "defaultValue": "", "description": "Custom conditional access policy url." }] }, { "id": "5cf42378-d67d-4f36-ba46-e8b86229381d", "deletedDateTime": null, "displayName": "Password Rule Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide password rule settings.\n      ", "values": [{ "name": "LockoutDurationInSeconds", "type": "System.Int32", "defaultValue": "60", "description": "The duration in seconds of the initial lockout period." }, { "name": "LockoutThreshold", "type": "System.Int32", "defaultValue": "10", "description": "The number of failed login attempts before the first lockout period begins." }, { "name": "BannedPasswordList", "type": "System.String", "defaultValue": "", "description": "A tab-delimited banned password list." }, { "name": "EnableBannedPasswordCheck", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if the banned password check is turned on." }] }, { "id": "80661d51-be2f-4d46-9713-98a2fcaec5bc", "deletedDateTime": null, "displayName": "Prohibited Names Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names settings.\n      ", "values": [{ "name": "CustomBlockedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to block for application display names." }, { "name": "CustomBlockedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of reserved words to block for application display names." }] }, { "id": "aad3907d-1d1a-448b-b3ef-7bf7f63db63b", "deletedDateTime": null, "displayName": "Prohibited Names Restricted Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names restricted settings.\n      ", "values": [{ "name": "CustomAllowedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to allow for application display names." }, { "name": "CustomAllowedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of whole reserved words to allow for application display names." }, { "name": "DoNotValidateAgainstTrademark", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if prohibited names validation against trademark global list is disabled." }] }] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: '62375ab9-6b52-47ed-826b-58e47e0e304c' } } as any),
      new CommandError(`Resource '62375ab9-6b52-47ed-826b-58e47e0e304c' does not exist.`));
  });

  it('returns error when no template with the specified displayName found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates`) {
        return { "value": [{ "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }] }, { "id": "08d542b9-071f-4e16-94b0-74abb372e3d9", "deletedDateTime": null, "displayName": "Group.Unified.Guest", "description": "Settings for a specific Unified Group", "values": [{ "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in a specific Unified Group." }] }, { "id": "4bc7f740-180e-4586-adb6-38b2e9024e6b", "deletedDateTime": null, "displayName": "Application", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide application behavior.\n      ", "values": [{ "name": "EnableAccessCheckForPrivilegedApplicationUpdates", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if access check for application privileged updates is turned on." }] }, { "id": "898f1161-d651-43d1-805c-3b0b388a9fc2", "deletedDateTime": null, "displayName": "Custom Policy Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide custom policy settings.\n      ", "values": [{ "name": "CustomConditionalAccessPolicyUrl", "type": "System.String", "defaultValue": "", "description": "Custom conditional access policy url." }] }, { "id": "5cf42378-d67d-4f36-ba46-e8b86229381d", "deletedDateTime": null, "displayName": "Password Rule Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide password rule settings.\n      ", "values": [{ "name": "LockoutDurationInSeconds", "type": "System.Int32", "defaultValue": "60", "description": "The duration in seconds of the initial lockout period." }, { "name": "LockoutThreshold", "type": "System.Int32", "defaultValue": "10", "description": "The number of failed login attempts before the first lockout period begins." }, { "name": "BannedPasswordList", "type": "System.String", "defaultValue": "", "description": "A tab-delimited banned password list." }, { "name": "EnableBannedPasswordCheck", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if the banned password check is turned on." }] }, { "id": "80661d51-be2f-4d46-9713-98a2fcaec5bc", "deletedDateTime": null, "displayName": "Prohibited Names Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names settings.\n      ", "values": [{ "name": "CustomBlockedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to block for application display names." }, { "name": "CustomBlockedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of reserved words to block for application display names." }] }, { "id": "aad3907d-1d1a-448b-b3ef-7bf7f63db63b", "deletedDateTime": null, "displayName": "Prohibited Names Restricted Settings", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for managing tenant-wide prohibited names restricted settings.\n      ", "values": [{ "name": "CustomAllowedSubStringsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of substring reserved words to allow for application display names." }, { "name": "CustomAllowedWholeWordsList", "type": "System.String", "defaultValue": "", "description": "A comma delimited list of whole reserved words to allow for application display names." }, { "name": "DoNotValidateAgainstTrademark", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if prohibited names validation against trademark global list is disabled." }] }] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { displayName: 'Invalid' } } as any),
      new CommandError(`Resource 'Invalid' does not exist.`));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { displayName: 'Invalid' } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if neither the id nor the displayName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both the id and the displayName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22', displayName: 'Group.Unified' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the displayName is specified', async () => {
    const actual = await command.validate({ options: { displayName: 'Group.Unified' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
