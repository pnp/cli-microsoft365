import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./sitescript-list');

describe(commands.SITESCRIPT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITESCRIPT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists available site scripts', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false } });
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
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false } });
  });

  it('correctly handles OData error when creating site script', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    await assert.rejects(command.action(logger, { options: { debug: false } } as any), new CommandError('An error has occurred'));
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