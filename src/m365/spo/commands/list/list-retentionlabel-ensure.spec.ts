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
import command, { options } from './list-retentionlabel-ensure.js';

describe(commands.LIST_RETENTIONLABEL_ENSURE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()! as typeof options;
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_RETENTIONLABEL_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should handle error when trying to set label', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
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
        return { "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        name: 'abc',
        listTitle: 'MyLibrary'
      })
    }), new CommandError("Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"));
  });

  it('should handle error if list does not exist', async () => {

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        throw new Error("404 - \"404 FILE NOT FOUND\"");
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        name: 'abc'
      })
    }), new CommandError('404 - "404 FILE NOT FOUND"'));
  });

  it('should set label for list with deprecated label', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return { "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        name: 'abc'
      })
    });
    const lastCall = postStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
    assert.strictEqual(lastCall.data.complianceTagValue, 'abc');
    assert.strictEqual(lastCall.data.syncToItems, false);
  });

  it('should set label for list (debug)', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists/getByTitle('MyLibrary')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return { "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listTitle: 'MyLibrary',
        name: 'abc'
      })
    });
    const lastCall = postStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
    assert.strictEqual(lastCall.data.complianceTagValue, 'abc');
    assert.strictEqual(lastCall.data.syncToItems, false);
  });

  it('should set label for list using listId (debug)', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team1/_api/web/lists(guid'4d535433-2a7b-40b0-9dad-8f0f8f3b3841')/?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return { "RootFolder": { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "MyLibrary", "ProgID": null, "ServerRelativeUrl": "/sites/team1/MyLibrary", "TimeCreated": "2019-01-11T10:03:19Z", "TimeLastModified": "2019-01-11T10:03:20Z", "UniqueId": "faaa6af2-0157-4e9a-a352-6165195923c8", "WelcomePage": "" } };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listId: '4d535433-2a7b-40b0-9dad-8f0f8f3b3841',
        name: 'abc'
      })
    });
    const lastCall = postStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
    assert.strictEqual(lastCall.data.complianceTagValue, 'abc');
    assert.strictEqual(lastCall.data.syncToItems, false);
  });

  it('should set label for list using blockDelete,blockEdit,syncToItems options', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team1/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/team1',
        listUrl: 'MyLibrary',
        name: 'abc',
        syncToItems: true
      })
    });
    const lastCall = postStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.listUrl, 'https://contoso.sharepoint.com/sites/team1/MyLibrary');
    assert.strictEqual(lastCall.data.complianceTagValue, 'abc');
    assert.strictEqual(lastCall.data.blockDelete, false);
    assert.strictEqual(lastCall.data.blockEdit, false);
    assert.strictEqual(lastCall.data.syncToItems, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', name: 'abc' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'foo is not a valid SharePoint Online site URL.');
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listid option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc', listId: 'XXXXX' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'XXXXX is not a valid GUID');
  });

  it('passes validation if the listid option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if listId, listUrl and listTitle options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc', listId: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});