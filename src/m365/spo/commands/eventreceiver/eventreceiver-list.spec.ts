import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./eventreceiver-list');

describe(commands.EVENTRECEIVER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const eventReceiverResponseJson = [
    {
      "ReceiverAssembly": "Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
      "ReceiverClass": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "ReceiverId": "625b1f4c-2869-457f-8b41-bed72059bb2b",
      "ReceiverName": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "SequenceNumber": 10000,
      "Synchronization": 1,
      "EventType": 309,
      "ReceiverUrl": null
    },
    {
      "ReceiverAssembly": "Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
      "ReceiverClass": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "ReceiverId": "41ad359e-ac6a-4a5e-8966-a85492ca4f52",
      "ReceiverName": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "SequenceNumber": 10000,
      "Synchronization": 1,
      "EventType": 310,
      "ReceiverUrl": null
    }
  ];

  const eventReceiverValue = {
    value: eventReceiverResponseJson
  };

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EVENTRECEIVER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['ReceiverId', 'ReceiverName']);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is set to site and one of the list properties is filled in', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', scope: 'site', listTitle: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the list ID is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list id', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list title', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list url', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if title and id are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is invalid value', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', scope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly handles list not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers`) > -1) {
        throw {
          error: {
            "odata.error": {
              "code": "-1, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } } as any),
      new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."));
  });

  it('retrieves all web event receivers', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all site event receivers', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all list event receivers queried by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all list event receivers queried by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all list event receivers queried by guid', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });
});
