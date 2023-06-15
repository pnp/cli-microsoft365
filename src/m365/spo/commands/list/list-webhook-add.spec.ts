import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-webhook-add');

describe(commands.LIST_WEBHOOK_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_WEBHOOK_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uses correct API url when list id option is passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listId: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
  });

  it('uses correct API url when list title option is passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
  });

  it('adds a webhook by passing list id (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0987cfd9-f02c-479b-9fb4-3f0550462848')/Subscriptions`) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: '0987cfd9-f02c-479b-9fb4-3f0550462848',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list url', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions`) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listUrl: '/sites/ninja/lists/Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list title including a client state', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        return {
          'clientState': 'awesome state',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        clientState: 'awesome state'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'awesome state',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list title including a expiration date', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-01-09T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2019-01-09'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-01-09T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('correctly handles a random API error', async () => {
    const errorMessage = 'An error has occurred';
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: errorMessage
          }
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2019-01-09'
      }
    } as any), new CommandError(errorMessage));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'foo',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    }, commandInfo);
    assert.strictEqual(typeof (actual), 'string');
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the list id option is not a valid GUID', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '12345',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    }, commandInfo);
    assert.strictEqual(typeof (actual), 'string');
  });

  it('passes validation if the listid option is a valid GUID', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime is in the past', async () => {
    const currentDate: Date = new Date();
    const currentMonth: number = currentDate.getMonth() + 1;
    const dateString: string = `${(currentDate.getFullYear() - 1)}-${currentMonth < 10 ? '0' : ''}${currentMonth}-01`;

    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: dateString
      }
    }, commandInfo);
    assert.strictEqual(actual, 'Provide an expiration date which is a date time in the future and within 6 months from now');
  });

  it('fails validation if the expirationDateTime more than six months from now', async () => {
    const currentDate: Date = new Date();
    const currentMonth: number = currentDate.getMonth() + 1;
    const dateString: string = `${(currentDate.getFullYear() + 1)}-${currentMonth < 10 ? '0' : ''}${currentMonth}-01`;

    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: dateString
      }
    }, commandInfo);
    assert.strictEqual(actual, 'Provide an expiration date which is a date time in the future and within 6 months from now');
  });

  it('passes validation if the expirationDateTime is in the future but no more than six months from now', async () => {
    const currentDate: Date = new Date();
    currentDate.setMonth(currentDate.getMonth() + 4);
    currentDate.setDate(1);
    const dateString: string = currentDate.toISOString().substr(0, 10);

    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: dateString
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date string', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-X-09'
      }
    }, commandInfo);
    assert.strictEqual(typeof (actual), 'string');
  });

  it('fails validation if the expirationDateTime option is not a valid date string (json output)', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-X-09',
        output: 'json'
      }
    }, commandInfo);
    assert.strictEqual(typeof (actual), 'string');
  });
});
