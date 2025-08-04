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
import command from './homesite-set.js';
import { z } from 'zod';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.HOMESITE_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;
  let commandOptionsSchema: z.ZodTypeAny;
  const siteUrl = 'https://contoso.sharepoint.com/sites/Work';
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';

  const defaultResponse = {
    "value": `The Home site has been set to ${siteUrl}. It may take some time for the change to apply. Check aka.ms/homesites for details.`
  };

  const homeSiteCountResponse = {
    value: [
      { Url: 'https://contoso.sharepoint.com/sites/home1' }
    ]
  };

  const emptyHomeSiteCountResponse = {
    value: null
  };

  const multipleHomeSiteCountResponse = {
    value: [
      { Url: 'https://contoso.sharepoint.com/sites/home1' },
      { Url: 'https://contoso.sharepoint.com/sites/home2' }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uses SetSPHSite when home site count is 1 and only siteUrl is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSiteCountResponse; // 1 home site
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/SetSPHSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
      sphSiteUrl: siteUrl
    });
  });

  it('sets the specified site as the Home Site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/SetSPHSite`) {
        return defaultResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        verbose: true
      }
    } as any);
    assert(loggerLogSpy.calledWith());
  });

  it('sets the specified site as the Home Site with vivaConnectionsDefaultStart using SetSPHSiteWithConfiguration when home site is 1 or less', async () => {
    const requestBody = {
      sphSiteUrl: siteUrl,
      configuration: { vivaConnectionsDefaultStart: true }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSiteCountResponse;
        ;
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/SetSPHSiteWithConfiguration`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        vivaConnectionsDefaultStart: true
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('sets the specified site as the Home Site with vivaConnectionsDefaultStart using UpdateTargetedSite when multiple home sites', async () => {
    const requestBody = {
      siteUrl: siteUrl,
      configurationParam: { IsVivaConnectionsDefaultStartPresent: true, vivaConnectionsDefaultStart: true }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return multipleHomeSiteCountResponse;
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        vivaConnectionsDefaultStart: true
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('sets the specified site as the Home Site with draftMode', async () => {
    const requestBody = {
      siteUrl: siteUrl,
      configurationParam: { IsInDraftModePresent: true, isInDraftMode: true }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return emptyHomeSiteCountResponse;
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        draftMode: true
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('sets the specified site as the Home Site with targetedLicenseType to frontLineWorkers', async () => {
    const requestBody = {
      siteUrl: siteUrl,
      configurationParam: { IsTargetedLicenseTypePresent: true, TargetedLicenseType: 1 }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return multipleHomeSiteCountResponse;
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        targetedLicenseType: "frontLineWorkers"
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('sets the specified site as the Home Site with targetedLicenseType to informationWorkers', async () => {
    const requestBody = {
      siteUrl: siteUrl,
      configurationParam: { IsTargetedLicenseTypePresent: true, TargetedLicenseType: 2 }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return multipleHomeSiteCountResponse;
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        targetedLicenseType: "informationWorkers"
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('covers transformAudienceNamesToIds with multiple audience names', async () => {
    const requestBody = {
      siteUrl: siteUrl,
      configurationParam: {
        IsAudiencesPresent: true,
        Audiences: ['00000000-0000-0000-0000-000000000001', '00000000-0000-0000-0000-000000000002']
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return multipleHomeSiteCountResponse;
      }
      return 'Invalid request';
    });

    // Mock entraGroup.getGroupIdByDisplayName to return different IDs for different names
    const entraGroupStub = sinon.stub(entraGroup, 'getGroupIdByDisplayName');
    entraGroupStub.withArgs('Marketing Team').resolves('00000000-0000-0000-0000-000000000001');
    entraGroupStub.withArgs('Sales Team').resolves('00000000-0000-0000-0000-000000000002');

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        audienceNames: 'Marketing Team, Sales Team'
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
    assert(entraGroupStub.calledWith('Marketing Team'));
    assert(entraGroupStub.calledWith('Sales Team'));
    assert.strictEqual(entraGroupStub.callCount, 2);

    entraGroupStub.restore();
  });

  it('uses SetSPHSite when home site count is 0 and only siteUrl is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return emptyHomeSiteCountResponse; // 0 home sites
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/SetSPHSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
      sphSiteUrl: siteUrl
    });
  });

  it('sets the specified site as the Home Site with multiple configuration options', async () => {
    const requestBody = {
      siteUrl: siteUrl,
      configurationParam: {
        IsAudiencesPresent: true,
        IsInDraftModePresent: true,
        IsVivaConnectionsDefaultStartPresent: true,
        IsOrderPresent: true,
        IsTargetedLicenseTypePresent: true,
        Order: 1,
        TargetedLicenseType: 0,
        isInDraftMode: false,
        vivaConnectionsDefaultStart: true,
        Audiences: ['00000000-0000-0000-0000-000000000001']
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return multipleHomeSiteCountResponse;
      }
      return 'Invalid request';
    });

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        vivaConnectionsDefaultStart: true,
        draftMode: false,
        audienceIds: '00000000-0000-0000-0000-000000000001',
        targetedLicenseType: "everyone",
        order: 1,
        verbose: true
      }
    } as any);

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles error when setting the Home Site', async () => {
    const errorResponse = {
      error: {
        "odata.error": {
          "code": "-2147213238, Microsoft.SharePoint.SPException",
          "message": {
            "lang": "en-US",
            "value": "The provided site url can't be set as a Home site."
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async () => {
      throw errorResponse;
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: siteUrl
      }
    } as any), new CommandError("The provided site url can't be set as a Home site."));
  });

  it('fails validation if the url is not a valid SharePoint url', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the siteUrl option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both audienceIds and audienceNames are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: {
        siteUrl: 'https://contoso.sharepoint.com',
        audienceIds: '00000000-0000-0000-0000-000000000001',
        audienceNames: 'Marketing Team'
      }
    });
    assert.strictEqual(actual.success, false);
  });

  it('correctly handles invalid GUIDs in audiences', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com', audienceIds: 'invalid-guid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if only audienceIds is specified', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com',
        audienceIds: '00000000-0000-0000-0000-000000000001'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only audienceNames is specified', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com',
        audienceNames: 'Marketing Team'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with valid targetedLicenseType values', async () => {
    const validTypes = ['everyone', 'frontLineWorkers', 'informationWorkers'];

    for (const type of validTypes) {
      const actual = await command.validate({
        options: {
          siteUrl: 'https://contoso.sharepoint.com',
          targetedLicenseType: type
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  });

  it('correctly handles non-integer order', async () => {
    const actual = commandOptionsSchema.safeParse({ siteUrl: 'https://contoso.sharepoint.com', order: -1 });
    assert.strictEqual(actual.success, false);
  });

  it('handles verbose mode correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/SetSPHSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: siteUrl,
        verbose: true
      }
    } as any);

    assert(log.some(entry => entry.includes('Setting the SharePoint home site')));
    assert(log.some(entry => entry.includes('Attempting to retrieve the SharePoint admin URL')));
  });
});