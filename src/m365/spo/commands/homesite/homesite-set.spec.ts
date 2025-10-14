import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './homesite-set.js';

describe(commands.HOMESITE_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const siteUrl = 'https://contoso.sharepoint.com/sites/Work';
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';

  const defaultResponse = {
    "value": [
      {
        "Audiences": [
          {
            "Email": "work@contoso.onmicrosoft.com",
            "Id": "7a1eea7f-9ab0-40ff-8f2e-0083d9d63451",
            "Title": "active Members"
          }
        ],
        "IsInDraftMode": true,
        "IsVivaBackendSite": false,
        "SiteId": "431d7819-4aaf-49a1-b664-b2fe9e609b63",
        "TargetedLicenseType": 2,
        "Title": "Work",
        "Url": "https://contoso.sharepoint.com/sites/Work",
        "VivaConnectionsDefaultStart": true,
        "WebId": "626c1724-8ac8-45d5-af87-c07c752fab75"
      }
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
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly sets the option vivaConnectionsDefaultStart', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        url: siteUrl,
        vivaConnectionsDefaultStart: true
      }
    });

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
      siteUrl: siteUrl,
      configurationParam: { IsVivaConnectionsDefaultStartPresent: true, vivaConnectionsDefaultStart: true }
    });
  });

  it('correctly sets the option draftMode', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        url: siteUrl,
        draftMode: true
      }
    });

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
      siteUrl: siteUrl,
      configurationParam: { IsInDraftModePresent: true, isInDraftMode: true }
    });
  });

  it('correctly sets the option targetedLicenseType to frontLineWorkers', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        url: siteUrl,
        targetedLicenseType: "frontLineWorkers"
      }
    });

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
      siteUrl: siteUrl,
      configurationParam: { IsTargetedLicenseTypePresent: true, TargetedLicenseType: 1 }
    });
  });

  it('correctly sets the option targetedLicenseType to informationWorkers', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        url: siteUrl,
        targetedLicenseType: "informationWorkers"
      }
    });

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
      siteUrl: siteUrl,
      configurationParam: { IsTargetedLicenseTypePresent: true, TargetedLicenseType: 2 }
    });
  });

  it('correctly sets the option audienceNames with multiple audience names', async () => {
    const entraGroupStub = sinon.stub(entraGroup, 'getGroupIdByDisplayName');
    entraGroupStub.withArgs('Marketing Team').resolves('00000000-0000-0000-0000-000000000001');
    entraGroupStub.withArgs('Sales Team').resolves('00000000-0000-0000-0000-000000000002');

    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    try {
      await command.action(logger, {
        options: {
          url: siteUrl,
          audienceNames: 'Marketing Team, Sales Team'
        }
      });

      assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
        siteUrl: siteUrl,
        configurationParam: { IsAudiencesPresent: true, Audiences: ['00000000-0000-0000-0000-000000000001', '00000000-0000-0000-0000-000000000002'] }
      });
    }
    finally {
      entraGroupStub.restore();
    }
  });

  it('correctly clears audience names', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        url: siteUrl,
        audienceNames: ''
      }
    });

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
      siteUrl: siteUrl,
      configurationParam: { IsAudiencesPresent: true, Audiences: [] }
    });
  });

  it('correctly sets multiple configuration options', async () => {
    const postRequestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`) {
        return defaultResponse;
      }
      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        url: siteUrl,
        vivaConnectionsDefaultStart: true,
        draftMode: false,
        audienceIds: '00000000-0000-0000-0000-000000000001',
        targetedLicenseType: "everyone",
        order: 1,
        verbose: true
      }
    });

    assert.deepStrictEqual(postRequestStub.lastCall.args[0].data, {
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
    });
  });

  it('correctly handles errors', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw {
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
    });

    await assert.rejects(command.action(logger, {
      options: {
        url: siteUrl
      }
    }), new CommandError("The provided site url can't be set as a Home site."));
  });

  it('fails validation if the url is not a valid SharePoint url', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'invalid', audienceIds: '00000000-0000-0000-0000-000000000001' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', audienceIds: '00000000-0000-0000-0000-000000000001' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both audienceIds and audienceNames are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      options: {
        url: 'https://contoso.sharepoint.com',
        audienceIds: '00000000-0000-0000-0000-000000000001',
        audienceNames: 'Marketing Team'
      }
    });
    assert.strictEqual(actual.success, false);
  });

  it('correctly handles invalid GUIDs in audiences', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', audienceIds: 'invalid-guid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if only audienceIds is specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      url: 'https://contoso.sharepoint.com',
      audienceIds: '00000000-0000-0000-0000-000000000001'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if only audienceNames is specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      url: 'https://contoso.sharepoint.com',
      audienceNames: 'Marketing Team'
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly handles non-integer order', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', order: -1 });
    assert.strictEqual(actual.success, false);
  });
});