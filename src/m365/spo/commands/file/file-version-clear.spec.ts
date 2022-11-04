import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./file-version-clear');

describe(commands.FILE_VERSION_CLEAR, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  const validWebUrl = "https://contoso.sharepoint.com";
  const validFileUrl = "/Shared Documents/Document.docx";
  const validFileId = "7a9b8bb6-d5c4-4de9-ab76-5210a7879e89";

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_VERSION_CLEAR), true);
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

  it('prompts before removing all file history when confirm option not passed', async () => {

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        fileId: validFileId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing all file history when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
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
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(validFileUrl)}')/versions/DeleteAll()`) {
        return { statusCode: 200 };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        fileUrl: validFileUrl,
        confirm: true
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
        confirm: true
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('deletes all version history from a file with the fileUrl options asking to confirm', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(validFileUrl)}')/versions/DeleteAll()`) {
        return { statusCode: 200 };
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

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
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: validWebUrl,
        confirm: true
      }
    }), new CommandError(err));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});