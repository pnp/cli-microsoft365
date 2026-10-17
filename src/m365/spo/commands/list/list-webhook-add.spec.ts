import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './list-webhook-add.js';

describe(commands.LIST_WEBHOOK_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()! as typeof options;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        listId: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      })
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      })
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
      options: commandOptionsSchema.parse({
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: '0987cfd9-f02c-479b-9fb4-3f0550462848',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      })
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        verbose: true
      })
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
      options: commandOptionsSchema.parse({
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listUrl: '/sites/ninja/lists/Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      })
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        clientState: 'awesome state'
      })
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString().substring(0, 10)
      })
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString().substring(0, 10)
      })
    }), new CommandError(errorMessage));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo',
      listTitle: 'Documents',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo is not a valid SharePoint Online site URL.');
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the list id option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listId: '12345',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, '12345 is not a valid GUID');
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the expirationDateTime is in the past', () => {
    const currentDate: Date = new Date();
    const currentMonth: number = currentDate.getMonth() + 1;
    const dateString: string = `${(currentDate.getFullYear() - 1)}-${currentMonth < 10 ? '0' : ''}${currentMonth}-01`;

    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listTitle: 'Documents',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: dateString
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Provide an expiration date which is a date time in the future and within 6 months from now');
  });

  it('fails validation if the expirationDateTime more than six months from now', () => {
    const currentDate: Date = new Date();
    const currentMonth: number = currentDate.getMonth() + 1;
    const dateString: string = `${(currentDate.getFullYear() + 1)}-${currentMonth < 10 ? '0' : ''}${currentMonth}-01`;

    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listTitle: 'Documents',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: dateString
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Provide an expiration date which is a date time in the future and within 6 months from now');
  });

  it('passes validation if the expirationDateTime is in the future but no more than six months from now', () => {
    const currentDate: Date = new Date();
    currentDate.setMonth(currentDate.getMonth() + 4);
    currentDate.setDate(1);
    const dateString: string = currentDate.toISOString().substring(0, 10);

    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listTitle: 'Documents',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: dateString
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date string', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listTitle: 'Documents',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-X-09'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `Provide the date in one of the following formats:
          'YYYY-MM-DD'
          'YYYY-MM-DDThh:mm'
          'YYYY-MM-DDThh:mmZ'
          'YYYY-MM-DDThh:mm±hh:mm'`);
  });

  it('fails validation if the expirationDateTime option is not a valid date string (json output)', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listTitle: 'Documents',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-X-09',
      output: 'json'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      listTitle: 'Documents',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString().substring(0, 10),
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
