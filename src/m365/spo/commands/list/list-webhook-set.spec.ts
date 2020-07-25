import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-webhook-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_WEBHOOK_SET, () => {
  let log: any[];
  let cmdInstance: any;

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
    };
  });

  afterEach(() => {
    Utils.restore([
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.LIST_WEBHOOK_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uses correct API url when list id option is passed', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
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
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
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
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
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
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
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

  it('updates notification url and expiration date of the webhook by passing list title (debug)', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-10-09'
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates notification url and expiration date of the webhook by passing list id (verbose)', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-10-09'
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'cc27a922-8224-4296-90a5-ebbc54da2e77')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'cc27a922-8224-4296-90a5-ebbc54da2e77',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates notification url and expiration date of the webhook by passing list title', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-10-09'
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates notification url of the webhook by passing list title', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates expiration date of the webhook by passing list title', (done) => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      expirationDateTime: '2019-03-02'
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options:
      {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        expirationDateTime: '2019-03-02'
      }
    }, () => {
      try {
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({
      options:
      {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        expirationDateTime: '2019-03-02'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if both list id and title options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webhook id option is not passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'foo',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '12345',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        id: '12345',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both notificationUrl and expirationDateTime options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e81'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the notificationUrl option is passed, but expirationDateTime is not', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if the expirationDateTime option is passed, but notificationUrl is not', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date string', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        expirationDateTime: '2018-X-09'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date string (json output)', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        expirationDateTime: '2018-X-09',
        output: 'json'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('supports verbose mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});