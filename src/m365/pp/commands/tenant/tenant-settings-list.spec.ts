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
const command: Command = require('./tenant-settings-list');

describe(commands.TENANT_SETTINGS_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TENANT_SETTINGS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['disableCapacityAllocationByEnvironmentAdmins', 'disableEnvironmentCreationByNonAdminUsers', 'disableNPSCommentsReachout', 'disablePortalsCreationByNonAdminUsers', 'disableSupportTicketsVisibleByAllUsers', 'disableSurveyFeedback', 'disableTrialEnvironmentCreationByNonAdminUsers', 'walkMeOptOut']);
  });

  it('successfully retrieves tenant settings', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/listtenantsettings?api-version=2020-10-01") {
        return successResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: {} } as any);
    assert(loggerLogSpy.calledWith(successResponse));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
