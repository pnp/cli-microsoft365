import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
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
import command from './list-webhook-list.js';

describe(commands.LIST_WEBHOOK_LIST, () => {
  const webhookListResponse = {
    value: [
      {
        "clientState": "pnp-js-core-subscription",
        "expirationDateTime": "2018-12-09T18:01:55.097Z",
        "id": "cfda40f2-6ca2-4424-9be0-33e9785b0e67",
        "notificationUrl": "https://deletemetestfunction.azurewebsites.net/api/FakeWebhookEndpoint?code=QlM2zaeJRti4WFGQUEqSo1ZmKMtRdB2JQ3mc2kzPj2aX6pNBAWVU4w==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      },
      {
        "clientState": '',
        "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
        "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
        "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
        "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
        "resourceData": null
      }
    ]
  };
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_WEBHOOK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'clientState', 'expirationDateTime', 'resource']);
  });

  it('retrieves all webhooks of the specific list if listTitle option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        listTitle: 'Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(webhookListResponse.value));
  });

  it('retrieves all webhooks of the specific list if listId option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(webhookListResponse.value));
  });

  it('retrieves all webhooks of the specific list if listUrl option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        listUrl: '/sites/ninja/lists/Documents',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith(webhookListResponse.value));
  });

  it('renders empty string for clientState, if no value for clientState was specified in the webhook', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogSpy.calledWith(webhookListResponse.value));
  });

  it('outputs user-friendly message when no webhooks found in verbose mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            "value": []
          };
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    });
    assert(loggerLogToStderrSpy.calledWith('No webhooks found'));
  });

  it('outputs all properties when output is JSON', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookListResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        output: 'json'
      }
    });
    assert(loggerLogSpy.calledWith(webhookListResponse.value));
  });

  it('command correctly handles list get reject request', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return {
          FormDigestValue: 'abc'
        };
      }

      throw 'Invalid request';
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    const actionTitle: string = 'Documents';

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        listTitle: actionTitle,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(err));
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint url', async () => {
    const actual = await command.validate({ options: { webUrl: 'notavalidurl', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCC' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });
});
