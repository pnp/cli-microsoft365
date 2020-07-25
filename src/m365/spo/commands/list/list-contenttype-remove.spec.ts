import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-contenttype-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_CONTENTTYPE_REMOVE, () => {
  let log: any[];
  let cmdInstance: any;
  let requests: any[];
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    requests = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.LIST_CONTENTTYPE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing content type from list when confirmation argument not passed (listId)', (done) => {
    cmdInstance.action({
      options: {
        debug: false,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A'
      }
    }, () => {
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

  it('prompts before removing content type from list when confirmation argument not passed (listTitle)', (done) => {
    cmdInstance.action({
      options: {
        debug: false,
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A'
      }
    }, () => {
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

  it('aborts removing content type from list when prompt not confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/lists(guid'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({
      options: {
        debug: false,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A'
      }
    }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes content type from list when listId option is passed and prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    const listId: string = 'dfddade1-4729-428d-881e-7fedf3cae50d';
    const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

    cmdInstance.action({
      options: {
        debug: true,
        listId: listId,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: contentTypeId
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'${encodeURIComponent(listId)}')/ContentTypes('${encodeURIComponent(contentTypeId)}')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['If-Match'] === '*') {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes content type from list when listTitle option is passed and prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    const listTitle: string = 'Documents';
    const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

    cmdInstance.action({
      options: {
        debug: true,
        listTitle: listTitle,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: contentTypeId
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('${encodeURIComponent(listTitle)}')/ContentTypes('${encodeURIComponent(contentTypeId)}')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['If-Match'] === '*') {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes content type from list when listId option is passed and prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    const listId: string = 'dfddade1-4729-428d-881e-7fedf3cae50d';
    const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

    cmdInstance.action({
      options: {
        debug: false,
        listId: listId,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: contentTypeId
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'${encodeURIComponent(listId)}')/ContentTypes('${encodeURIComponent(contentTypeId)}')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['If-Match'] === '*') {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes content type from list when listTitle option is passed and prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    const listTitle: string = 'Documents';
    const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

    cmdInstance.action({
      options: {
        debug: false,
        listTitle: listTitle,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: contentTypeId
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('${encodeURIComponent(listTitle)}')/ContentTypes('${encodeURIComponent(contentTypeId)}')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['If-Match'] === '*') {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles list get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const listTitle: string = 'Documents';
    const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

    cmdInstance.action({
      options: {
        debug: true,
        listTitle: listTitle,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: contentTypeId,
        confirm: true
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when listTitle option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const listTitle: string = 'Documents';
    const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

    cmdInstance.action({
      options: {
        debug: false,
        listTitle: listTitle,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: contentTypeId,
        confirm: true
      }
    }, () => {
      try {
        assert(1 === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when listId option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const listId: string = 'dfddade1-4729-428d-881e-7fedf3cae50d';
    const contentTypeId: string = '0x010109010053EE7AEB1FC54A41B4D9F66ADBDC312A';

    cmdInstance.action({
      options: {
        debug: false,
        listId: listId,
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: contentTypeId,
        confirm: true
      }
    }, () => {
      try {
        assert(1 === 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if both listId and listTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } });
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } });
    assert(actual);
  });

  it('passes validation if the listTitle option is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', contentTypeId: '0x0120' } });
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', contentTypeId: '0x0120' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the contentTypeId option is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures contentTypeId as string option', () => {
    const types = (command.types() as CommandTypes);
    ['c', 'contentTypeId'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });
});