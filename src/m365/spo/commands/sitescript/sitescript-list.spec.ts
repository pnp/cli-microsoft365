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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './sitescript-list.js';

describe(commands.SITESCRIPT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITESCRIPT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists available site scripts', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return {
          value: [
            {
              Content: null,
              Description: "description",
              Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
              Title: "script1",
              Version: 1
            },
            {
              Content: null,
              Description: "Contoso theme script description",
              Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
              Title: "Contoso theme script",
              Version: 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([
      {
        Content: null,
        Description: "description",
        Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
        Title: "script1",
        Version: 1
      },
      {
        Content: null,
        Description: "Contoso theme script description",
        Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
        Title: "Contoso theme script",
        Version: 1
      }
    ]));
  });

  it('lists available site scripts (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return {
          value: [
            {
              Content: null,
              Description: "description",
              Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
              Title: "script1",
              Version: 1
            },
            {
              Content: null,
              Description: "Contoso theme script description",
              Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
              Title: "Contoso theme script",
              Version: 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        Content: null,
        Description: "description",
        Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
        Title: "script1",
        Version: 1
      },
      {
        Content: null,
        Description: "Contoso theme script description",
        Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
        Title: "Contoso theme script",
        Version: 1
      }
    ]));
  });

  it('correctly handles no available site scripts', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWithExactly([]));
  });

  it('correctly handles OData error when creating site script', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
