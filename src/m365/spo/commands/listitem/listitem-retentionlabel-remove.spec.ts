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
import command, { options } from './listitem-retentionlabel-remove.js';

describe(commands.LISTITEM_RETENTIONLABEL_REMOVE, () => {

  //#region Mock Responses
  const listDetailsMock = {
    "RootFolder": {
      "ServerRelativeUrl": "/sites/project-x/list"
    }
  };
  //#endregion

  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listUrl = 'sites/project-x/list';
  const listTitle = 'test';
  const listId = 'b2307a39-e878-458b-bc90-03bc578531d6';

  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_RETENTIONLABEL_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing retentionlabel when confirmation argument not passed (id)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ listItemId: 1, webUrl: webUrl, listTitle: listTitle }) });

    assert(promptIssued);
  });

  it('aborts removing list item when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1
      })
    });
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('removes the retentionlabel based on listId when prompt confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listDetailsMock;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"","itemIds":[1]}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listId: listId,
        webUrl: webUrl,
        listItemId: 1
      })
    }));
  });

  it('removes the retentionlabel based on listTitle when prompt confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listDetailsMock;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"","itemIds":[1]}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1
      })
    }));
  });

  it('removes the retentionlabel based on listUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"","itemIds":[1]}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        force: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      })
    }));
  });

  it('removes the retentionlabel based on listUrl when prompt confirmed (debug)', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"","itemIds":[1]}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      })
    }));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'post').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        force: true,
        listUrl: listUrl,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 1
      })
    }), new CommandError(errorMessage));
  });

  it('fails validation if listId, listTitle and listUrl are not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listItemId: 1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listItemId: 1, listTitle: listTitle });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 1 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: '12345', listItemId: 1 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, '12345 is not a valid GUID');
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 1 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both listId and listTitle options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listTitle: listTitle, listItemId: 1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listItemId is not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listItemId is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listItemId: 'abc', listTitle: listTitle });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 1, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});