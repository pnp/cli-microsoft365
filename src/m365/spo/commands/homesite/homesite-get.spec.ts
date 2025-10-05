import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './homesite-get.js';

describe(commands.HOMESITE_GET, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';

  const homeSiteResult = {
    Audiences: [],
    IsInDraftMode: false,
    IsVivaBackendSite: false,
    SiteId: '53ad95dc-5d2c-42a3-a63c-716f7b8014f5',
    TargetedLicenseType: 0,
    Title: 'Work @ Contoso',
    Url: 'https://contoso.sharepoint.com/sites/Work',
    VivaConnectionsDefaultStart: false,
    WebId: '288ce497-483c-4cd5-b8a2-27b726d002e2'
  };

  const homeSiteResponse = [
    {
      Audiences: [],
      IsInDraftMode: true,
      IsVivaBackendSite: false,
      SiteId: '1f2a3b4c-5d6e-4789-a0b1-2c3d4e5f6a7b',
      TargetedLicenseType: 1,
      Title: 'Marketing',
      Url: 'https://contoso.sharepoint.com/sites/marketing/',
      VivaConnectionsDefaultStart: true,
      WebId: 'e1f2d3c4-b5a6-4d7e-8f90-1a2b3c4d5e6f'
    },
    {
      Audiences: [],
      IsInDraftMode: false,
      IsVivaBackendSite: true,
      SiteId: '9b8a7c6d-5e4f-4a3b-92c1-0dfe12345678',
      TargetedLicenseType: 2,
      Title: 'Operations',
      Url: 'https://contoso.sharepoint.com/sites/operations/',
      VivaConnectionsDefaultStart: false,
      WebId: '0a1b2c3d-4e5f-40a6-b1c2-3d4e5f6a7b8c'
    },
    homeSiteResult
  ];

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site url', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/Home' });
    assert.strictEqual(actual.success, true);
  });

  it('gets information about a Home Site', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSiteResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: homeSiteResult.Url, verbose: true } });
    assert(loggerLogSpy.calledWith(homeSiteResult));
  });

  it('gets information about a Home Site with trailing slash', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSiteResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: homeSiteResult.Url + '/', verbose: true } });
    assert(loggerLogSpy.calledWith(homeSiteResult));
  });

  it('gets information about a Home Site with capitalized URL', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSiteResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: homeSiteResult.Url.toUpperCase(), verbose: true } });
    assert(loggerLogSpy.calledWith(homeSiteResult));
  });

  it('outputs error when home site does not exist', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSiteResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/nonexistent' } }),
      new CommandError(`Home site with URL 'https://contoso.sharepoint.com/sites/nonexistent' not found.`));
  });

  it('correctly handles OData error when retrieving available home sites', async () => {
    sinon.stub(odata, 'getAllItems').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred.' } } } });

    await assert.rejects(command.action(logger, { options: { url: homeSiteResult.Url } }),
      new CommandError('An error has occurred.'));
  });
});
