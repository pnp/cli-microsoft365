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
const command: Command = require('./eventreceiver-get');

describe(commands.EVENTRECEIVER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const eventReceiverResponseJson = {
    "ReceiverAssembly": "",
    "ReceiverClass": "",
    "ReceiverId": "c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec",
    "ReceiverName": "PnP Test Receiver",
    "SequenceNumber": 30846,
    "Synchronization": 1,
    "EventType": 1,
    "ReceiverUrl": "https://pnp.github.io"
  };

  const eventReceiverValue = {
    value: [eventReceiverResponseJson]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    sinonUtil.restore([
      appInsights.trackEvent,
      pid.getProcessName,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.EVENTRECEIVER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['name', 'id']]);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com', name: 'Event receiver' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is set to site and one of the list properties is filled in', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', scope: 'site', listTitle: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the list ID is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with id', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with name', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list id', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list title', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list url', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if title and id are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is invalid value', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', scope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('throws error when multiple eventreceivers with the same name were found', async () => {
    const multipleEventreceiversResponse = {
      value: [
        { ReceiverId: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { ReceiverId: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleEventreceiversResponse;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' }
    }), new CommandError(`Multiple event receivers with name 'PnP Test Receiver' found: ${multipleEventreceiversResponse.value.map(x => x.ReceiverId).join(',')}`));
  });

  it('throws error when no eventreceiver with name were found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ value: [] });
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' } } as any),
      new CommandError(`The specified event receiver 'PnP Test Receiver' does not exist.`));
  });

  it('correctly handles list not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
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

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' } } as any),
      new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."));
  });

  it('throws error when event receiver with ID does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('Documents')/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) {
        return { 'odata.null': true };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        listTitle: 'Documents',
        id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec'
      }
    } as any), new CommandError(`The specified event receiver 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' does not exist.`));
  });

  it('retrieves web event receiver using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'PnP Test Receiver' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves site event receiver using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', name: 'PnP Test Receiver' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receiver retrieved by list title using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by url using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents', name: 'PnP Test Receiver' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by list id using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3', name: 'PnP Test Receiver' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves web event receiver using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves site event receiver using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec'
      }
    });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receiver retrieved by list title using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by url using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by list id using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' } });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
}); 