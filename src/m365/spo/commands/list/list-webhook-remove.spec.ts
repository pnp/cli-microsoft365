import commands from '../../commands';
import Command, { CommandValidate, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-webhook-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_WEBHOOK_REMOVE, () => {
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
      request.delete
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
    assert.strictEqual(command.name.startsWith(commands.LIST_WEBHOOK_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing webhook from list when confirmation argument not passed (list title)', (done) => {
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, () => {
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

  it('prompts before removing list when confirmation argument not passed (list id)', (done) => {
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, () => {
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

  it('aborts removing list when prompt not confirmed', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the list when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
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
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the list when prompt confirmed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
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
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('uses correct API url when list id option is passed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listId: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        confirm: true
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when list title option is passed', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        confirm: true
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if both list id and title options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webhook id option is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } });
    assert(actual);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
    assert(actual);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
    assert.notStrictEqual(actual, true);
  });
});