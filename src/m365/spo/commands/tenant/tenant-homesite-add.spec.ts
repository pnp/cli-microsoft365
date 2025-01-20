import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-homesite-add.js';

describe(commands.TENANT_HOMESITE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const homeSite = "https://contoso.sharepoint.com/sites/testcomms";
  const homeSites = {
    "Audiences": [],
    "IsInDraftMode": true,
    "IsVivaBackendSite": false,
    "SiteId": "ca49054c-85f3-41eb-a290-46ffda8f219c",
    "TargetedLicenseType": 0,
    "Title": "testcommsite",
    "Url": homeSite,
    "VivaConnectionsDefaultStart": false,
    "WebId": "256c4f0f-e372-47b4-a891-b4888e829e20"
  };

  const homeSiteConfig = {
    "Audiences": [
      {
        "Email": "SharingTest@reshmeeauckloo.onmicrosoft.com",
        "Id": "af8c0bc8-7b1b-44b4-b087-ffcc8df70d16",
        "Title": "SharingTest Members"
      }
    ],
    "IsInDraftMode": true,
    "IsVivaBackendSite": false,
    "SiteId": "ca49054c-85f3-41eb-a290-46ffda8f219c",
    "TargetedLicenseType": 0,
    "Title": "testcommsite",
    "Url": "https://contoso.sharepoint.com/sites/testcomms",
    "VivaConnectionsDefaultStart": false,
    "WebId": "256c4f0f-e372-47b4-a891-b4888e829e20"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    commandInfo = cli.getCommandInfo(command);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_HOMESITE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly logs command response', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPHSite/AddHomeSite`) {
        return homeSites;
      }

      throw opts.url;
    });

    await command.action(logger, { options: { url: homeSite, verbose: true } });
    assert(loggerLogSpy.calledWith(homeSites));
  });

  it('adds a home site with the specified URL, isInDraftMode, vivaConnectionsDefaultStart, and audiences', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPHSite/AddHomeSite`) {
        return homeSiteConfig;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        url: homeSite,
        isInDraftMode: 'true',
        vivaConnectionsDefaultStart: 'false',
        audiences: 'af8c0bc8-7b1b-44b4-b087-ffcc8df70d16',
        order: 2
      }
    });

    const expectedData = {
      "audiences": [
        "af8c0bc8-7b1b-44b4-b087-ffcc8df70d16"
      ],
      "isInDraftMode": "true",
      "order": 2,
      "siteUrl": "https://contoso.sharepoint.com/sites/testcomms",
      "vivaConnectionsDefaultStart": "false"
    };
    assert.deepStrictEqual(postStub.lastCall.args[0].data, expectedData);
  });

  it('fails validation if the url is not a valid SharePoint url', async () => {
    const actual = await command.validate({
      options: {
        url: "test"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with URL', async () => {
    const actual = await command.validate({
      options: {
        url: homeSite
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with URL and optional parameters', async () => {
    const actual = await command.validate({
      options: {
        url: homeSite,
        isInDraftMode: 'true',
        vivaConnectionsDefaultStart: 'false',
        audiences: 'af8c0bc8-7b1b-44b4-b087-ffcc8df70d16',
        order: 2
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles OData error when adding a home site', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, { options: { url: 'https://....' } } as any), new CommandError('An error has occurred'));
  });
});
