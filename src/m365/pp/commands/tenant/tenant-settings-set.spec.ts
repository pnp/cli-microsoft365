import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-settings-set.js';

describe(commands.TENANT_SETTINGS_SET, () => {
  let commandInfo: CommandInfo;

  const successResponse = {
    walkMeOptOut: false,
    disableNPSCommentsReachout: false,
    disableNewsletterSendout: false,
    disableEnvironmentCreationByNonAdminUsers: false,
    disablePortalsCreationByNonAdminUsers: false,
    disableSurveyFeedback: false,
    disableTrialEnvironmentCreationByNonAdminUsers: false,
    disableCapacityAllocationByEnvironmentAdmins: false,
    disableSupportTicketsVisibleByAllUsers: false,
    powerPlatform: {
      search: {
        disableDocsSearch: false,
        disableCommunitySearch: false,
        disableBingVideoSearch: false
      },
      teamsIntegration: {
        shareWithColleaguesUserLimit: 10000
      },
      powerApps: {
        disableShareWithEveryone: false,
        enableGuestsToMake: false,
        disableMembersIndicator: false,
        disableMakerMatch: false
      },
      environments: {
        disablePreferredDataLocationForTeamsEnvironment: false
      },
      governance: {
        disableAdminDigest: true,
        disableDeveloperEnvironmentCreationByNonAdminUsers: false
      },
      licensing: {
        disableBillingPolicyCreationByNonAdminUsers: false,
        storageCapacityConsumptionWarningThreshold: 85
      },
      powerPages: {},
      champions: {
        disableChampionsInvitationReachout: false,
        disableSkillsMatchInvitationReachout: false
      },
      intelligence: {
        disableCopilot: false,
        enableOpenAiBotPublishing: false
      },
      modelExperimentation: {
        enableModelDataSharing: false
      }
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if not one property is specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the shareWithColleaguesUserLimit is not a valid number', async () => {
    const actual = await command.validate({ options: { shareWithColleaguesUserLimit: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the shareWithColleaguesUserLimit is a negative number', async () => {
    const actual = await command.validate({ options: { shareWithColleaguesUserLimit: -1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the shareWithColleaguesUserLimit is a float number', async () => {
    const actual = await command.validate({ options: { shareWithColleaguesUserLimit: 3.14 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the shareWithColleaguesUserLimit is a valid number', async () => {
    const actual = await command.validate({ options: { shareWithColleaguesUserLimit: 9 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the storageCapacityConsumptionWarningThreshold is not a valid number', async () => {
    const actual = await command.validate({ options: { storageCapacityConsumptionWarningThreshold: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageCapacityConsumptionWarningThreshold is a negative number', async () => {
    const actual = await command.validate({ options: { storageCapacityConsumptionWarningThreshold: -1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the storageCapacityConsumptionWarningThreshold is a float number', async () => {
    const actual = await command.validate({ options: { storageCapacityConsumptionWarningThreshold: 3.14 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the storageCapacityConsumptionWarningThreshold is a valid number', async () => {
    const actual = await command.validate({ options: { storageCapacityConsumptionWarningThreshold: 9 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('successfully updates tenant settings', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/updateTenantSettings?api-version=2020-10-01") {
        return successResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        walkMeOptOut: false,
        disableNPSCommentsReachout: false,
        disableNewsletterSendout: false,
        disableEnvironmentCreationByNonAdminUsers: false,
        disablePortalsCreationByNonAdminUsers: false,
        disableSurveyFeedback: false,
        disableTrialEnvironmentCreationByNonAdminUsers: false,
        disableCapacityAllocationByEnvironmentAdmins: false,
        disableSupportTicketsVisibleByAllUsers: false,
        disableDocsSearch: false,
        disableCommunitySearch: false,
        disableBingVideoSearch: false,
        shareWithColleaguesUserLimit: 10000,
        disableShareWithEveryone: false,
        enableGuestsToMake: false,
        disableMembersIndicator: false,
        disableMakerMatch: false,
        disablePreferredDataLocationForTeamsEnvironment: false,
        disableAdminDigest: false,
        disableDeveloperEnvironmentCreationByNonAdminUsers: false,
        disableBillingPolicyCreationByNonAdminUsers: false,
        storageCapacityConsumptionWarningThreshold: 85,
        disableChampionsInvitationReachout: false,
        disableSkillsMatchInvitationReachout: false,
        disableCopilot: false,
        enableOpenAiBotPublishing: false,
        enableModelDataSharing: false
      }
    } as any);
    assert(loggerLogSpy.calledWith(successResponse));
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        message: "The request content was invalid and could not be deserialized."
      }
    };

    sinon.stub(request, 'post').callsFake(async () => {
      throw error;
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError(error.error.message));
  });
});
