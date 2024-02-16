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
import command from './tenant-settings-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SETTINGS_LIST);
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
