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
import command, { options } from './eventreceiver-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.EVENTRECEIVER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EVENTRECEIVER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'site.com', name: 'Event receiver' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'site.com is not a valid SharePoint Online site URL.');
  });

  it('fails validation if scope is set to site and one of the list properties is filled in', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', scope: 'site', listTitle: 'Documents' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the list ID is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listId: 'abc' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'abc is not a valid GUID.');
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'abc' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'abc is not a valid GUID.');
  });

  it('passes validation when all required parameters are valid with id', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid with name', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid and list id', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid and list title', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all required parameters are valid and list url', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listUrl: 'sites/hr-life/Lists/breakInheritance' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if title and id are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if title and id and url are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if title and url are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if scope is invalid value', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', scope: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Receiver', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('throws error when multiple eventreceivers with the same name were found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
      options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' })
    }), new CommandError("Multiple event receivers with name 'PnP Test Receiver' found. Found: 69703efe-4149-ed11-bba2-000d3adf7537, 3a081d91-5ea8-40a7-8ac9-abbaa3fcb893."));
  });

  it('handles selecting single result when multiple eventreceiver with the specified name found and cli is set to prompt', async () => {
    const multipleEventreceiversResponse = {
      value: [
        { ReceiverId: '69703efe-4149-ed11-bba2-000d3adf7537' },
        { ReceiverId: '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893' }
      ]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return multipleEventreceiversResponse;
        }
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(eventReceiverResponseJson);

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'PnP Test Receiver' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' }) }),
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' }) }),
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        listTitle: 'Documents',
        id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec'
      })
    }), new CommandError(`The specified event receiver 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' does not exist.`));
  });

  it('retrieves web event receiver using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'PnP Test Receiver' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves site event receiver using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', name: 'PnP Test Receiver' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receiver retrieved by list title using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by url using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents', name: 'PnP Test Receiver' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by list id using name as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers?$filter=receivername eq 'PnP Test Receiver'`) > -1) {
        return eventReceiverValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3', name: 'PnP Test Receiver' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves web event receiver using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' }) });
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        scope: 'site',
        id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec'
      })
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

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by url using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2FShared%20Documents')/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'Shared Documents', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });

  it('retrieves list event receivers queried by list id using id as option', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'b17bd74f-d1b1-42bf-a21d-f865a903acc3')/eventreceivers(guid'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec')`) > -1) {
        return eventReceiverResponseJson;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: 'b17bd74f-d1b1-42bf-a21d-f865a903acc3', id: 'c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec' }) });
    assert(loggerLogSpy.calledWith(eventReceiverResponseJson));
  });
}); 
