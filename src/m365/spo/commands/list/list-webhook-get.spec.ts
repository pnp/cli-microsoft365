import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-webhook-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_WEBHOOK_GET, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.LIST_WEBHOOK_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves specified webhook of the given list if title option is passed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "clientState": null,
            "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
            "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
            "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
            "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
            "resourceData": null
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "clientState": null,
          "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
          "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
          "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
          "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
          "resourceData": null
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves specified webhook of the given list if title option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "clientState": null,
            "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
            "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
            "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
            "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
            "resourceData": null
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "clientState": null,
          "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
          "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
          "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
          "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
          "resourceData": null
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if id option is passed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "clientState": null,
            "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
            "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
            "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
            "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
            "resourceData": null
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "clientState": null,
          "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
          "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
          "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
          "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
          "resourceData": null
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if id option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "clientState": null,
            "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
            "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
            "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
            "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
            "resourceData": null
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "clientState": null,
          "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
          "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
          "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
          "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
          "resourceData": null
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when getting information for a site that doesn\'t exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('ab27a922-8224-4296-90a5-ebbc54da1981')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ 
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'ab27a922-8224-4296-90a5-ebbc54da1981'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('404 - "404 FILE NOT FOUND"')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles list get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject(err);
    });

    const actionTitle: string = 'Documents';

    cmdInstance.action({
      options: {
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
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

  it('uses correct API url when id option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
        debug: false,
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

  it('fails validation if both list id and title options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webhook id option is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF'} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
    assert(actual);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } });
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
});