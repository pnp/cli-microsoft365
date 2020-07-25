import commands from '../../commands';
import Command, { CommandValidate, CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-webhook-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_WEBHOOK_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.LIST_WEBHOOK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves all webhooks of the specific list if title option is passed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if listTitle option is passed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if title option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if listTitle option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if id option is passed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if listId option is passed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if id option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webhooks of the specific list if listId option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: 'pnp-js-core-subscription',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('renders empty string for clientState, if no value for clientState was specified in the webhook', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": null,
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            id: 'cfda40f2-6ca2-4424-9be0-33e9785b0e67',
            clientState: '',
            expirationDateTime: '2018-12-09T18:01:55.097Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          },
          {
            id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
            clientState: '',
            expirationDateTime: '2019-01-27T16:32:05.4610008Z',
            resource: 'dfddade1-4729-428d-881e-7fedf3cae50d'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs user-friendly message when no webhooks found in verbose mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": []
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        verbose: true,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('No webhooks found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all properties when output is JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": "pnp-js-core-subscription",
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": null,
                "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
                "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
                "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        output: 'json'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [
            {
              "clientState": "pnp-js-core-subscription",
              "expirationDateTime": "2018-12-09T18:01:55.097Z",
              "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
              "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
              "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
              "resourceData": null
            }, {
              "clientState": null,
              "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
              "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
              "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
              "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
              "resourceData": null
            }
          ]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles list get reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionTitle: string = 'Documents';

    cmdInstance.action({
      options: {
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com',
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

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
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

  it('fails validation if both id and title options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', title: 'Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
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
});