import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./site-appcatalog-list');

describe(commands.SITE_APPCATALOG_LIST, () => {
  const appCatalogResponseValue = [
    {
      "AbsoluteUrl": "https://contoso.sharepoint.com/sites/site1",
      "ErrorMessage": null,
      "SiteID": "9798e615-b586-455e-8486-84913f492c49"
    },
    {
      "AbsoluteUrl": "https://contoso.sharepoint.com/sites/site2",
      "ErrorMessage": null,
      "SiteID": "686fe33a-7418-4a6b-92c9-d6170b1e3ae0"
    },
    {
      "AbsoluteUrl": "https://contoso.sharepoint.com/sites/site3",
      "ErrorMessage": "Success",
      "SiteID": "2f9fd04d-2674-40ca-9ad8-d7f982dce5d0"
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_APPCATALOG_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['AbsoluteUrl', 'SiteID']);
  });

  it('retrieves site collection app catalogs within the tenant (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return appCatalogResponseValue;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(appCatalogResponseValue));
  });

  it('correctly handles no site collection app catalogs in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return [];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert.strictEqual(log.length, 0);
  });

  it('correctly handles error when retrieving site collection app catalogs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        throw 'Invalid request';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Invalid request'));
  });
});