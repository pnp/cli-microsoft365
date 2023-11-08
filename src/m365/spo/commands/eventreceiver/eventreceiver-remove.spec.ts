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
import command from './eventreceiver-remove.js';

describe(commands.EVENTRECEIVER_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  const eventReceiverResponse = JSON.stringify(
    {
      "ReceiverAssembly": "Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
      "ReceiverClass": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "ReceiverId": "625b1f4c-2869-457f-8b41-bed72059bb2b",
      "ReceiverName": "Microsoft.SharePoint.Internal.SitePages.Sharing.PageSharingEventReceiver",
      "SequenceNumber": 10000,
      "Synchronization": 1,
      "EventType": 309,
      "ReceiverUrl": null
    }
  );

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
    (command as any).items = [];

    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.executeCommandWithOutput,
      Cli.promptForConfirmation,
      request.delete
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EVENTRECEIVER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', name: 'PnP Test Event Receiver' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list id and eventreceiver name is set', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', name: 'PnP Test Receiver' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list title', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', name: 'PnP Test Receiver' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list url and event receiver name', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listUrl: 'sites/hr-life/Lists/breakInheritance', name: 'PnP Test Receiver' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if list title and id are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', name: 'PnP Event Receiver' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if list id is invalid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: 'invalid', name: 'PnP Event Receiver' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if list id is filled in and scope is set to site', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', name: 'PnP Event Receiver', scope: 'site' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if list title is filled in and scope is set to site', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo list', name: 'PnP Event Receiver', scope: 'site' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if list url is filled in and scope is set to site', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listUrl: 'sites/hr-life/Lists/breakInheritance', name: 'PnP Event Receiver', scope: 'site' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance', name: 'PnP Event Receiver' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is invalid value', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', scope: 'abc', name: 'PnP Event Receiver' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is an invalid webUrl', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', name: 'PnP Event Receiver' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('prompts before removing the event receiver when force option not passed', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', name: 'PnP Test Receiver' } });

    assert(promptIssued);
  });

  it('aborts removing the event receiver when prompt not confirmed', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete');
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', name: 'PnP Test Receiver' } });
    assert(requestDeleteStub.notCalled);
  });

  it('deletes event receiver when prompt confirmed (debug)', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/site/eventreceivers('625b1f4c-2869-457f-8b41-bed72059bb2b')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(Cli, 'executeCommandWithOutput').resolves({
      stdout: eventReceiverResponse,
      stderr: ''
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', name: 'PnP Test Receiver', force: true } });
    assert(requestDeleteStub.called);
  });

  it('deletes event receiver with specified name', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/site/eventreceivers('625b1f4c-2869-457f-8b41-bed72059bb2b')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(Cli, 'executeCommandWithOutput').resolves({
      stdout: eventReceiverResponse,
      stderr: ''
    });
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', name: 'PnP Test Receiver' } });
    assert(requestDeleteStub.called);
  });

  it('deletes event receiver with by name from list retrieved by URL', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/GetList('${formatting.encodeQueryParameter('/sites/portal/Lists/rerlist')}')/eventreceivers('625b1f4c-2869-457f-8b41-bed72059bb2b')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(Cli, 'executeCommandWithOutput').resolves({
      stdout: eventReceiverResponse,
      stderr: ''
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'PnP Test Receiver', listUrl: '/sites/portal/Lists/rerlist' } });
    assert(requestDeleteStub.called);
  });

  it('deletes event receiver with by name from list retrieved by ID', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists('8fccab0d-78e5-4037-a6a7-0168f9359cd4')/eventreceivers('625b1f4c-2869-457f-8b41-bed72059bb2b')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(Cli, 'executeCommandWithOutput').resolves({
      stdout: eventReceiverResponse,
      stderr: ''
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', name: 'PnP Test Receiver', listId: '8fccab0d-78e5-4037-a6a7-0168f9359cd4' } });
    assert(requestDeleteStub.called);
  });

  it('deletes event receiver by specific id', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/site/eventreceivers('625b1f4c-2869-457f-8b41-bed72059bb2b')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(Cli, 'executeCommandWithOutput').resolves({
      stdout: eventReceiverResponse,
      stderr: ''
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', scope: 'site', id: '625b1f4c-2869-457f-8b41-bed72059bb2b' } });
    assert(requestDeleteStub.called);
  });

  it('deletes event receiver by specific name from specific list retrieved by the list title', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async opts => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle('${formatting.encodeQueryParameter('Documents')}')/eventreceivers('625b1f4c-2869-457f-8b41-bed72059bb2b')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(Cli, 'executeCommandWithOutput').resolves({
      stdout: eventReceiverResponse,
      stderr: ''
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents', name: 'PnP Test Receiver' } });
    assert(requestDeleteStub.called);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(Cli, 'executeCommandWithOutput').resolves({
      stdout: eventReceiverResponse,
      stderr: ''
    });

    sinon.stub(request, 'delete').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        scope: 'site',
        name: 'PnP Test Receiver',
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
