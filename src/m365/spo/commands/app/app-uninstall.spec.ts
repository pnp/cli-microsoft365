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
import command from './app-uninstall.js';

describe(commands.APP_UNINSTALL, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
    requests = [];
  });

  afterEach(() => {
    sinonUtil.restore(cli.promptForConfirmation);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_UNINSTALL);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uninstalls app from the specified site without prompting with confirmation argument (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('uninstalls app from the specified site without prompting with confirmation argument', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('uninstalls app from the specified site installed from the site collection app catalog', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true, appCatalogScope: 'sitecollection' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('prompts before uninstalling an app when confirmation argument not passed', async () => {
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } });
    assert(promptIssued);
  });

  it('aborts removing property when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } });
    assert(requests.length === 0);
  });

  it('uninstalls an app when prompt confirmed', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
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
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('correctly handles failure when app not found in app catalog', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
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

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
  });

  it('correctly handles failure when app is already being uninstalled', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw {
            error: JSON.stringify({
              'odata.error': {
                code: '-1, System.InvalidOperationException',
                message: {
                  value: 'Another job exists for this app instance. Please retry after that job is done.'
                }
              }
            })
          };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError('Another job exists for this app instance. Please retry after that job is done.'));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw { error: 'An error has occurred' };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (error message is not ODataError)', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw { error: JSON.stringify({ message: 'An error has occurred' }) };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError('{"message":"An error has occurred"}'));
  });

  it('correctly handles API OData error', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/uninstall`) > -1) {
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

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', force: true } } as any),
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

  it('passes validation when the scope is \'sitecollection\'', async () => {
    const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', siteUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'sitecollection' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
