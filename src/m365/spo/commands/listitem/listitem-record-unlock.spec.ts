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
import command, { options } from './listitem-record-unlock.js';

describe(commands.LISTITEM_RECORD_UNLOCK, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
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
    auth.connection.active = false;
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
      options: commandOptionsSchema.parse({
        verbose: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      })
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
      options: commandOptionsSchema.parse({
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1
      })
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
      options: commandOptionsSchema.parse({
        listId: listId,
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
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      })
    }), new CommandError(errorMessage));
  });

  it('fails validation if both id and title options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listItemId: 1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listItemId: 1, listTitle: listTitle });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 1 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: '12345', listItemId: 1 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, '12345 is not a valid GUID');
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 1 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if both id and title options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listTitle: listTitle, listItemId: 1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listItemId: 'abc', listTitle: listTitle });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listItemId: 1, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Unrecognized key: "unknownOption"');
  });
});