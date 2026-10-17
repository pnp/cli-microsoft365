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
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command, { options } from './listitem-isrecord.js';

describe(commands.LISTITEM_ISRECORD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-y';
  const listUrl = 'sites/project-x/documents';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const listIdResponse = { Id: '81f0ecee-75a8-46f0-b384-c8f4f9f31d99' };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const postFakes = async (opts: any): Promise<string> => {
    // requestObjectIdentity mock
    if (opts.data.indexOf('Name="Current"') > -1) {
      if ((opts.url as string).indexOf('returnerror.sharepoint.com') > -1) {
        await logger.log("Returns error from requestObjectIdentity");
        return Promise.reject("error occurred");
      }

      return Promise.resolve(JSON.stringify(
        [
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7618.1204",
            "ErrorInfo": null,
            "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
          },
          {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7",
            "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
          }
        ])
      );
    }

    // IsRecord request mocks
    if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
      // Unsuccessful response for when the item does not exist
      if ((opts.url as string).indexOf('itemdoesnotexist.sharepoint.com') > -1) {
        return Promise.resolve(JSON.stringify(
          [
            {
              "ErrorInfo": { "ErrorMessage": "Item does not exist. It may have been deleted by another user.", "ErrorValue": null, "TraceCorrelationId": "fedae69e-4077-8000-f13a-d4a607aefc32", "ErrorCode": -2130575338, "ErrorTypeName": "Microsoft.SharePoint.SPException" },
              "LibraryVersion": "16.0.9005.1214",
              "SchemaVersion": "15.0.0.0",
              "TraceCorrelationId": "fedae69e-4077-8000-f13a-d4a607aefc32"
            }]));
      }

      // Successful response
      return Promise.resolve(JSON.stringify(
        [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.9005.1214", "ErrorInfo": null, "TraceCorrelationId": "9ec8e69e-d001-8000-f13a-d5e03849cd96"
          }, 32, true
        ]
      ));
    }
    return Promise.reject('Invalid request');
  };

  const getFakes = async (opts: any): Promise<{ Id: string }> => {
    // Get list mock
    if ((opts.url as string).indexOf('/_api/web/lists') > -1 &&
      (opts.url as string).indexOf('$select=Id') > -1) {
      await logger.log('faked!');
      return Promise.resolve({
        Id: '81f0ecee-75a8-46f0-b384-c8f4f9f31d99'
      });
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-y/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`) {
      return Promise.resolve(listIdResponse);
    }
    return Promise.reject('Invalid request');
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ISRECORD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('throws an error when requesting a record for an item that does not exist', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Test List',
        id: 147,
        webUrl: `https://itemdoesnotexist.sharepoint.com/sites/project-y`,
        verbose: true
      })
    }), new CommandError('Item does not exist. It may have been deleted by another user.'));
  });

  it('test a record with list title passed in as an option', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listTitle: 'Test List',
        id: 147,
        webUrl: webUrl,
        verbose: true
      })
    });
    assert(loggerLogToStderrSpy.calledWith(`Getting list id for list Test List`));
  });

  it('test a record with list id passed in as an option', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
        id: 147,
        webUrl: webUrl,
        debug: true,
        verbose: true
      })
    });
    assert(loggerLogToStderrSpy.calledWith("List Id passed in as an argument."));
  });

  it('test a record with list url passed in as an option', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        listUrl: listUrl,
        id: 147,
        webUrl: webUrl,
        debug: true,
        verbose: true
      })
    });
    assert(loggerLogToStderrSpy.calledWith(`Getting list id for list ${listUrl}`));
  });

  it('fails to get _ObjecttIdentity_ when an error is returned by the _ObjectIdentity_ CSOM request', async () => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        listId: '99a14fe8-781c-3ce1-a1d5-c6e6a14561da',
        id: 147,
        webUrl: `https://returnerror.sharepoint.com/sites/project-y`
      })
    }), new CommandError('error occurred'));
    assert(loggerLogSpy.calledWith("Returns error from requestObjectIdentity"));
  });

  it('fails validation if listTitle, listId and listUrl option not specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: 1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listTitle, listId and listUrl are specified together', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', id: 1, listTitle: 'Test List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listUrl: listUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', listTitle: 'Test List', id: 1 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if the item ID is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Test List', id: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the item ID is not a positive number', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Test List', id: -1 });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and numerical ID specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listTitle: 'Test List', id: 1 });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: 'foo', id: 1 });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo in option listId is not a valid GUID');
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: 1, debug: true });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: 1, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Unrecognized key: "unknownOption"');
  });
});
