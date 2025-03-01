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
import command from './file-roleinheritance-reset.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.FILE_ROLEINHERITANCE_RESET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const fileUrl = '/sites/project-x/documents/Test1.docx';
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;
  const fileResponse = {
    CheckInComment: '',
    CheckOutType: 2,
    ContentTag: '{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12',
    CustomizedPageStatus: 0,
    ETag: '\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"',
    Exists: true,
    IrmEnabled: false,
    Length: '331673',
    Level: 1,
    LinkingUri: 'https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    LinkingUrl: 'https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866',
    ListItemAllFields: {
      Id: 1,
      ID: 1
    },
    MajorVersion: 3,
    MinorVersion: 0,
    Name: 'Test1.docx',
    ServerRelativeUrl: '/sites/project-x/documents/Test1.docx',
    TimeCreated: '2018-02-05T08:42:36Z',
    TimeLastModified: '2018-02-05T08:44:03Z',
    Title: '',
    UIVersion: 1536,
    UIVersionLabel: '3.0',
    UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      cli.promptForConfirmation,
      request.post,
      spo.getFileById
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_ROLEINHERITANCE_RESET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '0cd891ef-afce-4e55-b836-fce03286cccf', force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before resetting role inheritance for the file when force option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId
      }
    });


    assert(promptIssued);
  });

  it('aborts resetting role inheritance for the file when force option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId
      }
    });

    assert(postSpy.notCalled);
  });

  it('resets role inheritance on file by relative URL (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/resetroleinheritance`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        force: true
      }
    });
  });

  it('resets role inheritance on file by Id when prompt confirmed', async () => {
    sinon.stub(spo, 'getFileById').resolves(fileResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/resetroleinheritance`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fileId: fileId
      }
    });
  });

  it('correctly handles error when resetting file role inheritance', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        force: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
