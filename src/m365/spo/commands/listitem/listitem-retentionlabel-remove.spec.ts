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
import command from './listitem-retentionlabel-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

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
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
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
    await command.action(logger, { options: { listItemId: 1, webUrl: webUrl, listTitle: listTitle } });

    assert(promptIssued);
  });

  it('aborts removing list item when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: {
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1
      }
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
      options: {
        debug: true,
        listId: listId,
        webUrl: webUrl,
        listItemId: 1
      }
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
      options: {
        debug: true,
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1
      }
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
      options: {
        debug: true,
        force: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      }
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
      options: {
        debug: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      }
    }));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'post').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        force: true,
        listUrl: listUrl,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 1
      }
    }), new CommandError(errorMessage));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if both id and title options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listItemId: 1, listTitle: listTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 1 } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: '12345', listItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 1 } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both id and title options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listTitle: listTitle, listItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 'abc', listTitle: listTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});