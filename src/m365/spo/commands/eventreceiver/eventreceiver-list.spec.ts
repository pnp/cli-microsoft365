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
import command, { options } from './eventreceiver-list.js';

describe(commands.EVENTRECEIVER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'site.com' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `site.com is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if scope is set to site and one of the list properties is filled in', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', scope: 'site', listTitle: 'Documents' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the list ID is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: 'abc' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `abc is not a valid GUID.`);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid and list id', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid and list title', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid and list url', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', listUrl: 'sites/hr-life/Lists/breakInheritance' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if title and id are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if title and id and url are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if title and url are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if scope is invalid value', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', scope: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' }) }),
      new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."));
  });

  it('retrieves all web event receivers', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all site event receivers', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all list event receivers queried by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all list event receivers queried by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves all list event receivers queried by guid', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });
});
