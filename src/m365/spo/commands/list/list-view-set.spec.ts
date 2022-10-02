import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./list-view-set');

describe(commands.LIST_VIEW_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
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

  it('updates the Title of the list view specified using its name', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('List%201')/views/getByTitle('All%20items')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items', Title: 'All events' } });
    assert(loggerLogSpy.notCalled);
  });

  it('ignores global options when creating request data', async () => {
    const patchRequest: sinon.SinonStub = sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('List%201')/views/getByTitle('All%20items')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, verbose: false, output: "text", webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items', Title: 'All events' } });
    assert.deepEqual(patchRequest.lastCall.args[0].data, { Title: 'All events' });
  });

  it('updates the Title and CustomFormatter of the list view specified using its ID', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'330f29c5-5c4c-465f-9f4b-7903020ae1cf')/views/getById('330f29c5-5c4c-465f-9f4b-7903020ae1ce')`) {
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

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', Title: 'All events', CustomFormatter: 'abc' } });
  });

  it('correctly handles error when the specified list doesn\'t exist', async () => {
    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, System.ArgumentException",
            "message": {
              "lang": "en-US",
              "value": "List 'List' does not exist at site with URL 'https://contoso.sharepoint.com'."
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All items' } } as any),
      new CommandError("List 'List' does not exist at site with URL 'https://contoso.sharepoint.com'."));
  });

  it('correctly handles error when the specified view doesn\'t exist', async () => {
    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2147024809, System.ArgumentException",
            "message": {
              "lang": "en-US",
              "value": "The specified view is invalid."
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All items' } } as any),
      new CommandError("The specified view is invalid."));
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

  it('allows unknown options', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', listTitle: 'List 1', viewTitle: 'All items' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither listId nor listTitle specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', viewTitle: 'All items' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both listId and listTitle specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewTitle: 'All items' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId is not a GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'invalid', viewTitle: 'All items' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither viewId nor viewTitle specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both viewId and viewTitle specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewTitle: 'All items' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if viewId is not a GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when viewTitle and listTitle specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when viewId and listId specified and valid GUIDs', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});