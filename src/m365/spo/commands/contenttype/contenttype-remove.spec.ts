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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./contenttype-remove');

describe(commands.CONTENTTYPE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONTENTTYPE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('delete content type by id - prompt', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', confirm: false }
    } as any);
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });


  it('delete content type by id - prompt:continue', async () => {
    const postCallbackStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        id: '0x0100558D85B7216F6A489A499DB361E1AE2F',
        confirm: false
      }
    } as any);
    assert(postCallbackStub.called);
  });


  it('delete content type by id - prompt:declined', async () => {
    const postCallbackStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        id: '0x0100558D85B7216F6A489A499DB361E1AE2F',
        confirm: false
      }
    } as any);
    assert(postCallbackStub.notCalled);
  });

  it('delete content type by name - prompt', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'TestContentType')`) > -1) {
        return Promise.resolve({ "value": [{ "Name": "TestContentType", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', confirm: false } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });


  it('delete content type by name - prompt:continue', async () => {
    const getCallbackStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'TestContentType')`) > -1) {
        return Promise.resolve({ "value": [{ "Name": "TestContentType", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] });
      }

      return Promise.reject('Invalid request');
    });

    const postCallbackStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { debug: true, verbose: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', confirm: false } });
    assert(getCallbackStub.called);
    assert(postCallbackStub.called);
  });

  it('delete content type by name - prompt:declined', async () => {
    const postCallbackStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'TestContentType')`) > -1) {
        return Promise.resolve({ "value": [{ "Name": "TestContentType", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', confirm: false } });
    assert(postCallbackStub.notCalled);
  });

  it('correctly escapes special characters in the content type name', async () => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'Test%20Content%20Type')`) > -1) {
        return Promise.resolve({ "value": [{ "Name": "Test Content Type", "StringId": "0x0100558D85B7216F6A489A499DB361E1AE2F" }] });
      }

      return Promise.reject('Invalid request');
    });

    const postStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'Test Content Type', confirm: true } } as any);
    assert(getStub.called);
    assert(postStub.called);
  });


  it('correctly handles site content type not found by id', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', confirm: true } } as any),
      new CommandError('Content type not found'));
  });

  it('correctly handles site content type not found by name', async () => {
    //NonExistentContentType
    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/availableContentTypes?$filter=(Name eq 'NonExistentContentType')`) > -1) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject('Invalid request');
    });

    const deleteRequestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'NonExistentContentType', confirm: true } } as any),
      new CommandError('Content type not found'));

    assert(getRequestStub.called);
    assert(deleteRequestStub.notCalled);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'NonExistentContentType', confirm: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures id as string option', () => {
    const types = command.types;
    ['i', 'id'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('supports verbose mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the content type ID nor content type Name parameters are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when contenttype id parameter is provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype name parameter is provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype id and confirm parameters are provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype name and confirm parameters are provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type', confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when neither name nor id are provided, but confirm is', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both name and id are provided', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
