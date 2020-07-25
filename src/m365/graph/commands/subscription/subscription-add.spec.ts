import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./subscription-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SUBSCRIPTION_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let mockNowNumber = Date.parse("2019-01-01T00:00:00.000Z");

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      Date.now
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
    assert.strictEqual(command.name.startsWith(commands.SUBSCRIPTION_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds subscription', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "me/mailFolders('Inbox')/messages",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2016-11-20T18:23:45.935Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        resource: "me/mailFolders('Inbox')/messages",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
      }
    }, () => {
      try {
        assert.strictEqual(JSON.stringify(log[0]), JSON.stringify({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "me/mailFolders('Inbox')/messages",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2016-11-20T18:23:45.935Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should use a resource (group) specific default expiration if no expirationDateTime is set (debug)', (done) => {
    sinon.stub(Date, 'now').callsFake(() => mockNowNumber);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2019-01-03T22:29:00.000Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        resource: "groups",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("Matching resource in default values 'groups' => 'groups'"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should use a resource (group) specific default expiration if no expirationDateTime is set (verbose)', (done) => {
    sinon.stub(Date, 'now').callsFake(() => mockNowNumber);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2019-01-03T22:29:00.000Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        resource: "groups",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("An expiration maximum delay is resolved for the resource 'groups' : 4230 minutes."));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should use expirationDateTime if set (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2019-01-03T22:29:00.000Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        resource: "groups",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: "2019-01-03T00:00:00Z"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("Expiration date time is specified (2019-01-03T00:00:00Z)."));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should use a group specific default expiration if no expirationDateTime is set', (done) => {
    sinon.stub(Date, 'now').callsFake(() => mockNowNumber);
    let requestBodyArg: any = null;
    sinon.stub(request, 'post').callsFake((opts) => {
      requestBodyArg = opts.body;
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        resource: "groups",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    }, () => {
      try {
        // Expected for groups resource is 4230 minutes (-1 minutes for safe delay) = 72h - 1h31
        const expected = '2019-01-03T22:29:00.000Z';
        const actual = requestBodyArg.expirationDateTime;
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should use a generic default expiration if none can be found for the resource and no expirationDateTime is set (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          // NOTE Teams is not a supported resource and has no default maximum expiration delay
          "resource": "teams",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        // NOTE Teams is not a supported resource and has no default maximum expiration delay
        resource: "teams",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith("An expiration maximum delay couldn't be resolved for the resource 'teams'. Will use generic default value: 4230 minutes."));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should use a generic default expiration if none can be found for the resource and no expirationDateTime is set (debug)', (done) => {
    sinon.stub(Date, 'now').callsFake(() => mockNowNumber);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          // NOTE Teams is not a supported resource and has no default maximum expiration delay
          "resource": "teams",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        // NOTE Teams is not a supported resource and has no default maximum expiration delay
        resource: "teams",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
      }
    }, () => {
      try {
        // Expected for groups resource is 4230 minutes (-1 minutes for safe delay) = 72h - 1h31
        // assert.strictEqual(log[4], expectedLog);
        assert(cmdInstanceLogSpy.calledWith("Actual expiration date time: 2019-01-03T22:29:00.000Z"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        resource: "me/mailFolders('Inbox')/messages",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
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

  it('fails validation if expirationDateTime is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        resource: "me/mailFolders('Inbox')/messages",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: 'foo'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if notificationUrl is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        resource: "me/mailFolders('Inbox')/messages",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "foo",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if changeType is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        resource: "me/mailFolders('Inbox')/messages",
        changeType: 'foo',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientState exceeds maximum allowed length', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        resource: "me/mailFolders('Inbox')/messages",
        changeType: 'updated',
        clientState: 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: null
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the expirationDateTime is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        resource: "me/mailFolders('Inbox')/messages",
        changeType: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: null
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});