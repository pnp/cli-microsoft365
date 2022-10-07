import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
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
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_SETTINGS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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
