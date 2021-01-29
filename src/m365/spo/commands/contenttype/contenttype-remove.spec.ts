import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError, CommandTypes } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./contenttype-remove');

describe(commands.CONTENTTYPE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONTENTTYPE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('delete content type by id - prompt', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', confirm: false }
    } as any, (err?: any) => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('delete content type by id - prompt:continue', (done) => {
    const postCallbackStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        id: '0x0100558D85B7216F6A489A499DB361E1AE2F',
        confirm: false
      }
    } as any, (err?: any) => {

      try {
        assert(postCallbackStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('delete content type by id - prompt:declined', (done) => {
    const postCallbackStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, {
      options: {
        debug: true,
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        id: '0x0100558D85B7216F6A489A499DB361E1AE2F',
        confirm: false
      }
    } as any, (err?: any) => {
      try {
        assert(postCallbackStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('delete content type by name - prompt', (done) => {
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

    command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', confirm: false } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('delete content type by name - prompt:continue', (done) => {
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

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: true, verbose: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', confirm: false } }, () => {
      try {
        assert(getCallbackStub.called);
        assert(postCallbackStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('delete content type by name - prompt:declined', (done) => {
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

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'TestContentType', confirm: false } }, () => {
      try {
        assert(postCallbackStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes special characters in the content type name', (done) => {
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

    command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'Test Content Type', confirm: true } } as any, (err?: any) => {
      try {
        assert(getStub.called);
        assert(postStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('correctly handles site content type not found by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Content type not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles site content type not found by name', (done) => {
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

    command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'NonExistentContentType', confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Content type not found')));
        assert(getRequestStub.called);
        assert(deleteRequestStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => Promise.reject('An error has occurred'));

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'NonExistentContentType', confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures id as string option', () => {
    const types = (command.types() as CommandTypes);
    ['i', 'id'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports verbose mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = command.validate({ options: { webUrl: 'site.com', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the content type ID nor content type Name parameters are specified', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when contenttype id parameter is provided', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype name parameter is provided', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype id and confirm parameters are provided', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', confirm: true } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when contenttype name and confirm parameters are provided', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type', confirm: true } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when neither name nor id are provided, but confirm is', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', confirm: true } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both name and id are provided', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'Test Content Type', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } });
    assert.notStrictEqual(actual, true);
  });
});