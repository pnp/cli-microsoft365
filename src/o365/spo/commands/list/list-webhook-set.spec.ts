import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./list-webhook-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_WEBHOOK_SET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      },
    };
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.patch
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      request.patch
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LIST_WEBHOOK_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.LIST_WEBHOOK_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when list id option is passed', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
      if (opts.url.indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
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
        assert.equal(actual, expected);
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'cc27a922-8224-4296-90a5-ebbc54da2e77')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
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
        assert.equal(actual, expected);
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
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
        assert.equal(actual, expected);
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
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
        assert.equal(actual, expected);
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
      if (opts.url.indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.body);
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
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
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.equal(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.equal(actual, true);
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
    assert.equal(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.equal(actual, true);
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
    assert.equal(actual, true);
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
    assert.notEqual(actual, true);
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
    assert.notEqual(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.LIST_WEBHOOK_SET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});