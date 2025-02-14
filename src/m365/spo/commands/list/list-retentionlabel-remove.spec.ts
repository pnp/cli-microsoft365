import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-retentionlabel-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LIST_RETENTIONLABEL_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const listResponse = {
    "RootFolder": {
      "ServerRelativeUrl": "/sites/team1/Shared Documents"
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_RETENTIONLABEL_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the retentionlabel on the specified list when force option not passed (listTitle)', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary'
      }
    });

    assert(confirmationStub.calledOnce);
  });

  it('prompts before removing the retentionlabel on the specified list when force option not passed (listId)', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF'
      }
    });

    assert(confirmationStub.calledOnce);
  });

  it('prompts before removing the retentionlabel on the specified list when force option not passed (listUrl)', async () => {
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listUrl: '/sites/team1/MyLibrary'
      }
    });

    assert(confirmationStub.calledOnce);
  });

  it('aborts removing list retentionlabel when prompt not confirmed', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const getSpy = sinon.spy(request, 'get');
    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF'
      }
    });
    assert(getSpy.notCalled);
  });

  it('should handle error when trying to remove label', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) {
        throw {
          error: {
            'odata.error': {
              code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
              message: {
                value: 'Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        force: true
      }
    } as any), new CommandError("Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"));
  });

  it('should handle error if list does not exist', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: '404 - File not found'
          }
        }
      }
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        force: true
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });

  it('should remove label for list with listTitle (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        force: true
      }
    }));
  });

  it('should remove label for list with listId (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'faaa6af2-0157-4e9a-a352-6165195923c8')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: 'faaa6af2-0157-4e9a-a352-6165195923c8',
        force: true
      }
    }));
  });

  it('should remove label for list with listUrl (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listUrl: '/sites/team1/MyLibrary',
        force: true
      }
    }));
  });

  it('should remove label for list with listUrl when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listUrl: '/sites/team1/MyLibrary'
      }
    }));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listid option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'XXXXX' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listid option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if listId, listUrl and listTitle options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});