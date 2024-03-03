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
import command from './site-appcatalog-list.js';

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_APPCATALOG_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['AbsoluteUrl', 'SiteID']);
  });

  it('retrieves site collection app catalogs within the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return { value: appCatalogResponseValue };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(appCatalogResponseValue));
  });

  it('correctly handles error when retrieving site collection app catalogs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        throw { error: { error: { message: 'Something went wrong' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Something went wrong'));
  });
});