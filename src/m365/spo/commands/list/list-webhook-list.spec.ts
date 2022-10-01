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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-webhook-list');

describe(commands.LIST_WEBHOOK_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_WEBHOOK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'clientState', 'expirationDateTime', 'resource']);
  });

  it('retrieves all webhooks of the specific list if title option is passed (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: true,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('retrieves all webhooks of the specific list if listTitle option is passed (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: true,
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('retrieves all webhooks of the specific list if title option is passed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: false,
        title: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('retrieves all webhooks of the specific list if listTitle option is passed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: false,
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('retrieves all webhooks of the specific list if id option is passed (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: true,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('retrieves all webhooks of the specific list if listId option is passed (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: true,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('retrieves all webhooks of the specific list if id option is passed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: false,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('retrieves all webhooks of the specific list if listId option is passed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: false,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('renders empty string for clientState, if no value for clientState was specified in the webhook', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": [
              {
                "clientState": '',
                "expirationDateTime": "2018-12-09T18:01:55.097Z",
                "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
                "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
                "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
                "resourceData": null
              }, {
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: false,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "clientState": '',
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }, {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]));
  });

  it('outputs user-friendly message when no webhooks found in verbose mode', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            "value": []
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogToStderrSpy.calledWith('No webhooks found'));
  });

  it('outputs all properties when output is JSON', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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
                "clientState": '',
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

    await command.action(logger, {
      options: {
        debug: false,
        id: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        output: 'json'
      }
    });
    assert(loggerLogSpy.calledWith(
      [
        {
          "clientState": "pnp-js-core-subscription",
          "expirationDateTime": "2018-12-09T18:01:55.097Z",
          "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
          "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
          "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
          "resourceData": null
        }, {
          "clientState": '',
          "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
          "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
          "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
          "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
          "resourceData": null
        }
      ]
    ));
  });

  it('command correctly handles list get reject request', async () => {
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

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(err));
  });

  it('uses correct API url when id option is passed', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await command.action(logger, {
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
  });

  it('fails validation if both id and title options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both id and title options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', title: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both listId and listTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
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