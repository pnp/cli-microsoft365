import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-retract.js';

describe(commands.APP_RETRACT, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
    sinon.stub(request, 'get').resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_RETRACT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retracts app from the tenant app catalog (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('retracts app from the tenant app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('retracts app from the specified tenant app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps', force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`https://contoso.sharepoint.com/sites/apps/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('retracts app from the specified site collection app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com/sites/site1', appCatalogScope: 'sitecollection', force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`https://contoso.sharepoint.com/sites/site1/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('prompts before retracting an app when confirmation argument not passed', async () => {
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } });

    assert(promptIssued);
  });

  it('aborts retracting app when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } });
    assert(requests.length === 0);
  });

  it('retracts an app when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('correctly handles failure when app not found in app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw {
            error: JSON.stringify({
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                message: {
                  lang: "en-US",
                  value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                }
              }
            })
          };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw { error: 'An error has occurred' };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (error message is not ODataError)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw { error: JSON.stringify({ message: 'An error has occurred' }) };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError('{"message":"An error has occurred"}'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/retract`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw {
            error: JSON.stringify({
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
                message: {
                  value: 'An error has occurred'
                }
              }
            })
          };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if the appCatalogUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is specified and the appCatalogUrl is not', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the id is not a GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when scope is \'sitecollection\', but no appCatalogUrl specified', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the scope is not \'tenant\' nor \'sitecollection\'', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id, appCatalogUrl and scope options are specified', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'sitecollection' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
