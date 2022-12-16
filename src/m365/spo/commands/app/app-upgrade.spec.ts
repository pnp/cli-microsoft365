import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-upgrade');

describe(commands.APP_UPGRADE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
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
    requests = [];
  });

  afterEach(() => {
    sinonUtil.restore(request.post);
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
    assert.strictEqual(command.name.startsWith(commands.APP_UPGRADE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('upgrades the app in the specified site (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('upgrades app in the specified site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('upgrades app in the specified site installed from site collection app catalog', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'sitecollection' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('correctly handles failure when app not found in app catalog', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.reject({
            error: JSON.stringify({
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                message: {
                  lang: "en-US",
                  value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                }
              }
            })
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } } as any),
      new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (error message is not ODataError)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.reject({ error: JSON.stringify({ message: 'An error has occurred' }) });
        }
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } } as any),
      new CommandError('{"message":"An error has occurred"}'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/upgrade`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.reject({
            error: JSON.stringify({
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
                message: {
                  value: 'An error has occurred'
                }
              }
            })
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation when the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123', siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the siteUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the scope is not \'tenant\' nor \'sitecollection\'', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id and siteUrl options are specified', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid id and site url', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is \'sitecollection\'', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'sitecollection' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
