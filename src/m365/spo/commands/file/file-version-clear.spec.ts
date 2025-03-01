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
import command from './file-version-clear.js';

describe(commands.FILE_VERSION_CLEAR, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;
  const validWebUrl = "https://contoso.sharepoint.com";
  const validFileUrl = "/Shared Documents/Document.docx";
  const validFileId = "7a9b8bb6-d5c4-4de9-ab76-5210a7879e89";

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
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_VERSION_CLEAR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if fileId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        fileId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileUrl: validFileUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (fileUrl)', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, fileUrl: validFileUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (fileId)', async () => {
    const actual = await command.validate({ options: { webUrl: validWebUrl, fileId: validFileId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing all file history when force option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        fileId: validFileId
      }
    });

    assert(promptIssued);
  });

  it('aborts removing all file history when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        fileId: validFileId
      }
    });
    assert(postSpy.notCalled);
  });

  it('deletes all version history from a file with the fileUrl options', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(validFileUrl)}')/versions/DeleteAll()`) {
        return { statusCode: 200 };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileUrl: validFileUrl,
        force: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('deletes all version history from a file with the fileId options', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileById('${validFileId}')/versions/DeleteAll()`) {
        return { statusCode: 200 };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileId: validFileId,
        force: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('deletes all version history from a file with the fileUrl options asking to confirm', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(validFileUrl)}')/versions/DeleteAll()`) {
        return { statusCode: 200 };
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileUrl: validFileUrl
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('deletes all version history from a file with the fileId options asking to confirm', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileById('${validFileId}')/versions/DeleteAll()`) {
        return { statusCode: 200 };
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileId: validFileId
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('command correctly handles version list reject request', async () => {
    const err = 'Invalid version request';
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: err
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        force: true
      }
    }), new CommandError(err));
  });
});
