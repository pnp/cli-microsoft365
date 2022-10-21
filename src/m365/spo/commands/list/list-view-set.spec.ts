import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { urlUtil } from '../../../../utils/urlUtil';
import { formatting } from '../../../../utils/formatting';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./list-view-set');

describe(commands.LIST_VIEW_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const listId = '0cd891ef-afce-4e55-b836-fce03286cccf';
  const listTitle = 'List 1';
  const listUrl = '/lists/List 1';
  const viewId = 'cc27a922-8224-4296-90a5-ebbc54da2e81';
  const viewTitle = 'All items';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: webUrl
    }));
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName,
      spo.getRequestDigest
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_VIEW_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('allows unknown options', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', listTitle: listTitle, viewTitle: viewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId is not a GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: 'invalid', viewTitle: viewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if viewId is not a GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: 'List 1', viewId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['listId', 'listTitle', 'listUrl'], ['viewId', 'viewTitle']]);
  });

  it('passes validation when viewId and listId specified as valid GUIDs', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, viewId: viewId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('ignores global options when creating request data', async () => {
    const patchRequest: sinon.SinonStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events' })) {
          return;
        }
      }

      return 'Invalid request';
    });

    await command.action(logger, { options: { debug: false, verbose: false, output: "text", webUrl: webUrl, listTitle: listTitle, viewTitle: viewTitle, Title: 'All events' } });
    assert.deepEqual(patchRequest.lastCall.args[0].data, { Title: 'All events' });
  });

  it('updates the Title of the list view specified using its name', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events' })) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: false, webUrl: webUrl, listTitle: listTitle, viewTitle: viewTitle, Title: 'All events' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates the Title and CustomFormatter of the list view specified using its ID', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/views/GetById('${formatting.encodeQueryParameter(viewId)}')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events', CustomFormatter: 'abc' })) {
          return;
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, webUrl: webUrl, listId: listId, viewId: viewId, Title: 'All events', CustomFormatter: 'abc' } });
  });

  it('updates the Title and CustomFormatter of the list view specified using its Url', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/views/GetById('${formatting.encodeQueryParameter(viewId)}')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events', CustomFormatter: 'abc' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, webUrl: webUrl, listUrl: listUrl, viewId: viewId, Title: 'All events', CustomFormatter: 'abc' } });
  });

  it('correctly handles error when updating existing list view', async () => {
    const errorMessage = 'request rejected';
    sinon.stub(request, 'patch').callsFake(async () => { throw errorMessage; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        webUrl: webUrl,
        listTitle: listTitle,
        viewTitle: viewTitle
      }
    }), new CommandError(errorMessage));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});