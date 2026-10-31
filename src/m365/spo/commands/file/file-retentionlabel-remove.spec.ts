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
import command, { options } from './file-retentionlabel-remove.js';

describe(commands.FILE_RETENTIONLABEL_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const fileUrl = `/Shared Documents/Fo'lde'r/Document.docx`;
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const listId = 1;
  const fileResponse = {
    ListItemAllFields: {
      Id: listId,
      ParentList: {
        RootFolder: {
          ServerRelativeUrl: "/Shared Documents"
        }
      }
    }
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean = false;

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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.promptForConfirmation,
      cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_RETENTIONLABEL_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing retentionlabel from a file when confirmation argument not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: webUrl, fileUrl: fileUrl }) });

    assert(promptIssued);
  });

  it('aborts removing file retention label when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        fileUrl: fileUrl,
        webUrl: webUrl
      })
    });
    assert(postSpy.notCalled);
  });

  it('removes the retentionlabel from a file based on fileUrl when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === `{"listUrl":"https://contoso.sharepoint.com/Shared Documents","complianceTagValue":"","itemIds":[1]}`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: commandOptionsSchema.parse({
        fileUrl: fileUrl,
        webUrl: webUrl
      })
    }));
  });

  it('removes the retentionlabel from a file based on fileId when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileById('${fileId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === `{"listUrl":"https://contoso.sharepoint.com/Shared Documents","complianceTagValue":"","itemIds":[1]}`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: commandOptionsSchema.parse({
        fileId: fileId,
        webUrl: webUrl
      })
    }));
  });

  it('removes the retentionlabel from a file based on fileId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileById('${fileId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === `{"listUrl":"https://contoso.sharepoint.com/Shared Documents","complianceTagValue":"","itemIds":[1]}`) {
        return;
      }

      throw 'Invalid request';
    });


    await assert.doesNotReject(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        fileId: fileId,
        webUrl: webUrl,
        force: true
      })
    }));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        force: true,
        fileUrl: fileUrl,
        webUrl: webUrl
      })
    }), new CommandError(errorMessage));
  });

  it('fails validation if both fileUrl or fileId options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', fileUrl: fileUrl });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileUrl: fileUrl });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the fileId option is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileId: '12345' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the fileId option is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileId: fileId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileId: fileId, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both fileId and fileUrl options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, fileId: fileId, fileUrl: fileUrl });
    assert.strictEqual(actual.success, false);
  });
});