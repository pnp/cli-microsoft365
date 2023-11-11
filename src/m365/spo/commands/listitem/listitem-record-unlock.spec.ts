import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './listitem-record-unlock.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LISTITEM_RECORD_UNLOCK, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const listUrl = "/MyLibrary";
  const listTitle = "MyLibrary";
  const listId = "cc27a922-8224-4296-90a5-ebbc54da2e85";
  const webUrl = "https://contoso.sharepoint.com";
  const listResponse = {
    "RootFolder": {
      "Exists": true,
      "IsWOPIEnabled": false,
      "ItemCount": 0,
      "Name": listTitle,
      "ProgID": null,
      "ServerRelativeUrl": listUrl,
      "TimeCreated": "2019-01-11T10:03:19Z",
      "TimeLastModified": "2019-01-11T10:03:20Z",
      "UniqueId": listId,
      "WelcomePage": ""
    }
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_RECORD_UNLOCK);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('unlocks a list item based on listUrl (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        verbose: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      }
    }));
  });

  it('unlocks a list item based on listTitle', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('${listTitle}')/?$expand=RootFolder&$select=RootFolder`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1
      }
    }));
  });

  it('unlocks a list item based on listId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'${listId}')/?$expand=RootFolder&$select=RootFolder`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        listId: listId,
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
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      }
    }), new CommandError(errorMessage));
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