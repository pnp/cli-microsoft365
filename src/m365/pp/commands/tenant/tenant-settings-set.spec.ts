import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { session } from '../../../../utils/session';
import { pid } from '../../../../utils/pid';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./tenant-settings-set');

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
    isableCapacityAllocationByEnvironmentAdmins: false,
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
        disableMembersIndicator: false
      },
      environments: {},
      governance: {
        disableAdminDigest: false
      },
      licensing: {
        disableBillingPolicyCreationByNonAdminUsers: false
      },
      powerPages: {}
    }
  };


  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
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
    assert.strictEqual(command.name, commands.TENANT_SETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if not one property is specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the walkMeOptOut option is set to false', async () => {
    const actual = await command.validate({ options: { walkMeOptOut: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the walkMeOptOut option is set to true', async () => {
    const actual = await command.validate({ options: { walkMeOptOut: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableNPSCommentsReachout option is set to false', async () => {
    const actual = await command.validate({ options: { disableNPSCommentsReachout: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableNPSCommentsReachout option is set to true', async () => {
    const actual = await command.validate({ options: { disableNPSCommentsReachout: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableNewsletterSendout option is set to false', async () => {
    const actual = await command.validate({ options: { disableNewsletterSendout: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableNewsletterSendout option is set to true', async () => {
    const actual = await command.validate({ options: { disableNewsletterSendout: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableEnvironmentCreationByNonAdminusers option is set to false', async () => {
    const actual = await command.validate({ options: { disableEnvironmentCreationByNonAdminusers: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableEnvironmentCreationByNonAdminusers option is set to true', async () => {
    const actual = await command.validate({ options: { disableEnvironmentCreationByNonAdminusers: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disablePortalsCreationByNonAdminusers option is set to false', async () => {
    const actual = await command.validate({ options: { disablePortalsCreationByNonAdminusers: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disablePortalsCreationByNonAdminusers option is set to true', async () => {
    const actual = await command.validate({ options: { disablePortalsCreationByNonAdminusers: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableSurveyFeedback option is set to false', async () => {
    const actual = await command.validate({ options: { disableSurveyFeedback: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableSurveyFeedback option is set to true', async () => {
    const actual = await command.validate({ options: { disableSurveyFeedback: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableTrialEnvironmentCreationByNonAdminusers option is set to false', async () => {
    const actual = await command.validate({ options: { disableTrialEnvironmentCreationByNonAdminusers: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableTrialEnvironmentCreationByNonAdminusers option is set to true', async () => {
    const actual = await command.validate({ options: { disableTrialEnvironmentCreationByNonAdminusers: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableCapacityAllocationByEnvironmentAdmins option is set to false', async () => {
    const actual = await command.validate({ options: { disableCapacityAllocationByEnvironmentAdmins: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableCapacityAllocationByEnvironmentAdmins option is set to true', async () => {
    const actual = await command.validate({ options: { disableCapacityAllocationByEnvironmentAdmins: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableSupportTicketsVisibleByAllUsers option is set to false', async () => {
    const actual = await command.validate({ options: { disableSupportTicketsVisibleByAllUsers: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableSupportTicketsVisibleByAllUsers option is set to true', async () => {
    const actual = await command.validate({ options: { disableSupportTicketsVisibleByAllUsers: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableDocsSearch option is set to false', async () => {
    const actual = await command.validate({ options: { disableDocsSearch: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableDocsSearch option is set to true', async () => {
    const actual = await command.validate({ options: { disableDocsSearch: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableCommunitySearch option is set to false', async () => {
    const actual = await command.validate({ options: { disableCommunitySearch: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableCommunitySearch option is set to true', async () => {
    const actual = await command.validate({ options: { disableCommunitySearch: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableBingVideoSearch option is set to false', async () => {
    const actual = await command.validate({ options: { disableBingVideoSearch: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableBingVideoSearch option is set to true', async () => {
    const actual = await command.validate({ options: { disableBingVideoSearch: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the shareWithColleaguesUserLimit is not a valid number', async () => {
    const actual = await command.validate({ options: { shareWithColleaguesUserLimit: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the shareWithColleaguesUserLimit is a valid number', async () => {
    const actual = await command.validate({ options: { shareWithColleaguesUserLimit: '9' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableShareWithEveryone option is set to false', async () => {
    const actual = await command.validate({ options: { disableShareWithEveryone: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableShareWithEveryone option is set to true', async () => {
    const actual = await command.validate({ options: { disableShareWithEveryone: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the enableGuestsToMake option is set to false', async () => {
    const actual = await command.validate({ options: { enableGuestsToMake: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the enableGuestsToMake option is set to true', async () => {
    const actual = await command.validate({ options: { enableGuestsToMake: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableAdminDigest option is set to false', async () => {
    const actual = await command.validate({ options: { disableAdminDigest: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableAdminDigest option is set to true', async () => {
    const actual = await command.validate({ options: { disableAdminDigest: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableDeveloperEnvironmentCreationByNonAdminUsers option is set to false', async () => {
    const actual = await command.validate({ options: { disableDeveloperEnvironmentCreationByNonAdminUsers: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableDeveloperEnvironmentCreationByNonAdminUsers option is set to true', async () => {
    const actual = await command.validate({ options: { disableDeveloperEnvironmentCreationByNonAdminUsers: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableBillingPolicyCreationByNonAdminUsers option is set to false', async () => {
    const actual = await command.validate({ options: { disableBillingPolicyCreationByNonAdminUsers: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableBillingPolicyCreationByNonAdminUsers option is set to true', async () => {
    const actual = await command.validate({ options: { disableBillingPolicyCreationByNonAdminUsers: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableChampionsInvitationReachout option is set to false', async () => {
    const actual = await command.validate({ options: { disableChampionsInvitationReachout: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableChampionsInvitationReachout option is set to true', async () => {
    const actual = await command.validate({ options: { disableChampionsInvitationReachout: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableSkillsMatchInvitationReachout option is set to false', async () => {
    const actual = await command.validate({ options: { disableSkillsMatchInvitationReachout: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the disableSkillsMatchInvitationReachout option is set to true', async () => {
    const actual = await command.validate({ options: { disableSkillsMatchInvitationReachout: true } }, commandInfo);
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
        disableEnvironmentCreationByNonAdminusers: false,
        disablePortalsCreationByNonAdminusers: false,
        disableSurveyFeedback: false,
        disableTrialEnvironmentCreationByNonAdminusers: false,
        disableCapacityAllocationByEnvironmentAdmins: false,
        disableSupportTicketsVisibleByAllUsers: false,
        disableDocsSearch: false,
        disableCommunitySearch: false,
        disableBingVideoSearch: false,
        shareWithColleaguesUserLimit: 10000,
        disableShareWithEveryone: false,
        enableGuestsToMake: false,
        disableAdminDigest: false,
        disableDeveloperEnvironmentCreationByNonAdminUsers: false,
        disableBillingPolicyCreationByNonAdminUsers: false,
        disableChampionsInvitationReachout: false,
        disableSkillsMatchInvitationReachout: false
      }
    } as any);
    assert(loggerLogSpy.calledWith(successResponse));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
