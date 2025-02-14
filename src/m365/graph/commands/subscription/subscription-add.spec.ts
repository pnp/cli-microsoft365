import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './subscription-add.js';

describe(commands.SUBSCRIPTION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const mockNowNumber = Date.parse("2019-01-01T00:00:00.000Z");

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Date.now
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SUBSCRIPTION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds subscription', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "me/mailFolders('Inbox')/messages",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2016-11-20T18:23:45.935Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
      }
    });
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
      "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
      "resource": "me/mailFolders('Inbox')/messages",
      "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
      "changeType": "updated",
      "clientState": "secretClientValue",
      "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
      "expirationDateTime": "2016-11-20T18:23:45.935Z",
      "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
    }));
  });

  it('should use a resource (group) specific default expiration if no expirationDateTime is set (debug)', async () => {
    sinon.stub(Date, 'now').returns(mockNowNumber);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2019-01-03T22:29:00.000Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        resource: "groups",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    });
    assert(loggerLogToStderrSpy.calledWith("Matching resource in default values 'groups' => 'groups'"));
  });

  it('should use a resource (group) specific default expiration if no expirationDateTime is set (verbose)', async () => {
    sinon.stub(Date, 'now').returns(mockNowNumber);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2019-01-03T22:29:00.000Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        resource: "groups",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    });
    assert(loggerLogToStderrSpy.calledWith("An expiration maximum delay is resolved for the resource 'groups' : 4230 minutes."));
  });

  it('should use expirationDateTime if set (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2019-01-03T22:29:00.000Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        resource: "groups",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: "2019-01-03T00:00:00Z"
      }
    });
    assert(loggerLogToStderrSpy.calledWith("Expiration date time is specified (2019-01-03T00:00:00Z)."));
  });

  it('should use a group specific default expiration if no expirationDateTime is set', async () => {
    sinon.stub(Date, 'now').returns(mockNowNumber);
    let requestBodyArg: any = null;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requestBodyArg = opts.data;
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "groups",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        resource: "groups",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    });
    // Expected for groups resource is 4230 minutes (-1 minutes for safe delay) = 72h - 1h31
    const expected = '2019-01-03T22:29:00.000Z';
    const actual = requestBodyArg.expirationDateTime;
    assert.strictEqual(actual, expected);
  });

  it('should use a generic default expiration if none can be found for the resource and no expirationDateTime is set (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          // NOTE Teams is not a supported resource and has no default maximum expiration delay
          "resource": "teams",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        // NOTE Teams is not a supported resource and has no default maximum expiration delay
        resource: "teams",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    });
    assert(loggerLogToStderrSpy.calledWith("An expiration maximum delay couldn't be resolved for the resource 'teams'. Will use generic default value: 4230 minutes."));
  });

  it('should use a generic default expiration if none can be found for the resource and no expirationDateTime is set (debug)', async () => {
    sinon.stub(Date, 'now').returns(mockNowNumber);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          // NOTE Teams is not a supported resource and has no default maximum expiration delay
          "resource": "teams",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        // NOTE Teams is not a supported resource and has no default maximum expiration delay
        resource: "teams",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient"
      }
    });
    // Expected for groups resource is 4230 minutes (-1 minutes for safe delay) = 72h - 1h31
    assert(loggerLogToStderrSpy.calledWith("Actual expiration date time: 2019-01-03T22:29:00.000Z"));
  });

  it('adds subscription and set specific TLS version', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "me/mailFolders('Inbox')/messages",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2016-11-20T18:23:45.935Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437",
          "latestSupportedTlsVersion": "v1_3"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        latestTLSVersion: 'v1_3'
      }
    });
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
      "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
      "resource": "me/mailFolders('Inbox')/messages",
      "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
      "changeType": "updated",
      "clientState": "secretClientValue",
      "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
      "expirationDateTime": "2016-11-20T18:23:45.935Z",
      "creatorId": "8ee44408-0679-472c-bc2a-692812af3437",
      "latestSupportedTlsVersion": "v1_3"
    }));
  });

  it('adds subscription and includes resource data', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/subscriptions`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
          "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
          "resource": "me/mailFolders('Inbox')/messages",
          "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
          "changeType": "updated",
          "clientState": "secretClientValue",
          "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
          "expirationDateTime": "2016-11-20T18:23:45.935Z",
          "creatorId": "8ee44408-0679-472c-bc2a-692812af3437",
          "includeResourceData": true,
          "encryptionCertificateId": "MyCert",
          "encryptionCertificate": "Q0xJIGZvciBNaWNyb3NvZnQgMzY1"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        includeResourceData: true,
        encryptionCertificate: 'Q0xJIGZvciBNaWNyb3NvZnQgMzY1',
        encryptionCertificateId: 'MyCert'
      }
    });
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscriptions/$entity",
      "id": "7f105c7d-2dc5-4530-97cd-4e7ae6534c07",
      "resource": "me/mailFolders('Inbox')/messages",
      "applicationId": "24d3b144-21ae-4080-943f-7067b395b913",
      "changeType": "updated",
      "clientState": "secretClientValue",
      "notificationUrl": "https://webhook.azurewebsites.net/api/send/myNotifyClient",
      "expirationDateTime": "2016-11-20T18:23:45.935Z",
      "creatorId": "8ee44408-0679-472c-bc2a-692812af3437",
      "includeResourceData": true,
      "encryptionCertificateId": "MyCert",
      "encryptionCertificate": "Q0xJIGZvciBNaWNyb3NvZnQgMzY1"
    }));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if expirationDateTime is not valid', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if notificationUrl is not valid', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "foo",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if changeTypes is not valid', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'foo',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientState exceeds maximum allowed length', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the notificationUrlAppId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        notificationUrlAppId: 'foo',
        expirationDateTime: null
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if latestTLSVersion is not valid', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        latestTLSVersion: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if lifecycleNotificationUrl is not valid', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: 'https://webhook.azurewebsites.net/api/send/myNotifyClient',
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        lifecycleNotificationUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if resource data should be included, but encryptionCertificate is not set', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: 'https://webhook.azurewebsites.net/api/send/myNotifyClient',
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        includeResourceData: true,
        encryptionCertificateId: 'MyCert'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if resource data should be included, but encryptionCertificateId is not set', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: 'https://webhook.azurewebsites.net/api/send/myNotifyClient',
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        includeResourceData: true,
        encryptionCertificate: 'Q0xJIGZvciBNaWNyb3NvZnQgMzY1'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the expirationDateTime is not specified', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the notificationUrl points to valid https URL', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the notificationUrl points to valid Azure Event Hub location', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "EventHub:https://exchangenotifications.servicebus.windows.net/eventhubname/inboxmessages?tenantId=contoso.com",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the notificationUrl points to valid Azure Event Grid Partner Topic', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "EventGrid:?azuresubscriptionid=b07a45b3-f7b7-489b-9269-da6f3f93dff0&resourcegroup=rg-graph-api&partnertopic=messages&location=germanywestcentral",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the lifecycleNotificationUrl points to valid https URL', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        lifecycleNotificationUrl: "https://webhook.azurewebsites.net/api/lifecycleNotifications",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the lifecycleNotificationUrl points to valid Azure Event Hub location', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "EventHub:https://exchangenotifications.servicebus.windows.net/eventhubname/inboxmessages?tenantId=contoso.com",
        lifecycleNotificationUrl: "EventHub:https://exchangenotifications.servicebus.windows.net/eventhubname/inboxmessages?tenantId=contoso.com",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the lifecycleNotificationUrl points to valid Azure Event Grid Partner Topic', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "EventGrid:?azuresubscriptionid=b07a45b3-f7b7-489b-9269-da6f3f93dff0&resourcegroup=rg-graph-api&partnertopic=messages&location=germanywestcentral",
        lifecycleNotificationUrl: "EventGrid:?azuresubscriptionid=b07a45b3-f7b7-489b-9269-da6f3f93dff0&resourcegroup=rg-graph-api&partnertopic=messages&location=germanywestcentral",
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the notificationUrlAppId is a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        notificationUrlAppId: '24d3b144-21ae-4080-943f-7067b395b913',
        expirationDateTime: null
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if latestTLSVersion is valid', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        latestTLSVersion: 'v1_3'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if resource data should be included and encryptionCertificate is specified together with encryptionCertificateId', async () => {
    const actual = await command.validate({
      options: {
        resource: "me/mailFolders('Inbox')/messages",
        changeTypes: 'updated',
        clientState: 'secretClientValue',
        notificationUrl: "https://webhook.azurewebsites.net/api/send/myNotifyClient",
        expirationDateTime: '2016-11-20T18:23:45.935Z',
        includeResourceData: true,
        encryptionCertificate: 'Q0xJIGZvciBNaWNyb3NvZnQgMzY1',
        encryptionCertificateId: 'MyCert'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
