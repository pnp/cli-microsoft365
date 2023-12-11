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
import command from './file-copy.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.FILE_COPY, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project';
  const documentName = 'Document.pdf';
  const relSourceUrl = '/sites/project/Documents/' + documentName;
  const absoluteSourceUrl = 'https://contoso.sharepoint.com/sites/project/Documents/' + documentName;
  const sourceId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const relTargetUrl = '/sites/project/Documents';
  const absoluteTargetUrl = 'https://contoso.sharepoint.com/sites/project/Documents';

  let log: any[];
  let logger: Logger;
  let requestPostStub: sinon.SinonStub;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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

    requestPostStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/SP.MoveCopyUtil.CopyFileByPath`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request URL: ' + opts.url;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_COPY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', sourceUrl: absoluteSourceUrl, targetUrl: absoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the sourceId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceId: 'invalid', targetUrl: absoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the sourceId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceId: sourceId, targetUrl: absoluteTargetUrl } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both sourceId and sourceUrl options are not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, targetUrl: absoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both sourceId and url options are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, sourceId: sourceId, sourceUrl: absoluteSourceUrl, targetUrl: absoluteTargetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if nameConflictBehavior has an invalid value', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceUrl: absoluteSourceUrl, targetUrl: absoluteTargetUrl, nameConflictBehavior: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required properties are provided', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceUrl: absoluteSourceUrl, targetUrl: absoluteTargetUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('copies file with sourceId successfully when absolute URLs are provided', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/GetFileById('${sourceId}')?$select=ServerRelativePath`) {
        return {
          ServerRelativePath: {
            DecodedUrl: absoluteTargetUrl + `/${documentName}`
          }
        };
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceId: sourceId,
        targetUrl: absoluteTargetUrl
      }
    });

    const response = {
      srcPath: {
        DecodedUrl: absoluteSourceUrl
      },
      destPath: {
        DecodedUrl: absoluteTargetUrl + `/${documentName}`
      },
      overwrite: false,
      options: {
        KeepBoth: false,
        ResetAuthorAndCreatedOnCopy: false,
        ShouldBypassSharedLocks: false
      }
    };

    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data, response);
  });

  it('copies file with absolute url successfully when absolute URLs are provided', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: absoluteSourceUrl,
        targetUrl: absoluteTargetUrl
      }
    });

    const response = {
      srcPath: {
        DecodedUrl: absoluteSourceUrl
      },
      destPath: {
        DecodedUrl: absoluteTargetUrl + `/${documentName}`
      },
      overwrite: false,
      options: {
        KeepBoth: false,
        ResetAuthorAndCreatedOnCopy: false,
        ShouldBypassSharedLocks: false
      }
    };

    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data, response);
  });

  it('copies file with relative url successfully when server relative URLs are provided', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl
      }
    });

    const response = {
      srcPath: {
        DecodedUrl: absoluteSourceUrl
      },
      destPath: {
        DecodedUrl: absoluteTargetUrl + `/${documentName}`
      },
      overwrite: false,
      options: {
        KeepBoth: false,
        ResetAuthorAndCreatedOnCopy: false,
        ShouldBypassSharedLocks: false
      }
    };

    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data, response);
  });

  it('copies file with relative url successfully with a new name', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        newName: 'Document-renamed.pdf'
      }
    });

    const response = {
      srcPath: {
        DecodedUrl: absoluteSourceUrl
      },
      destPath: {
        DecodedUrl: absoluteTargetUrl + '/Document-renamed.pdf'
      },
      overwrite: false,
      options: {
        KeepBoth: false,
        ResetAuthorAndCreatedOnCopy: false,
        ShouldBypassSharedLocks: false
      }
    };

    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data, response);
  });

  it('copies file with relative url successfully with nameConflictBehavior fail', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        nameConflictBehavior: 'fail'
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.overwrite, false, 'Overwrite option is not false');
    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.KeepBoth, false, 'KeepBoth option is not false');
  });

  it('copies file with relative url successfully with nameConflictBehavior replace', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        nameConflictBehavior: 'replace'
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.overwrite, true, 'Overwrite option is not true');
    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.KeepBoth, false, 'KeepBoth option is not false');
  });

  it('copies file with relative url successfully with nameConflictBehavior rename', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        nameConflictBehavior: 'rename'
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.overwrite, false, 'Overwrite option is not false');
    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.KeepBoth, true, 'KeepBoth option is not true');
  });

  it('copies file with relative url successfully with bypassSharedLock option', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        bypassSharedLock: true
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.ShouldBypassSharedLocks, true);
  });

  it('copies file with relative url successfully with resetAuthorAndCreated option', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        resetAuthorAndCreated: true
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.ResetAuthorAndCreatedOnCopy, true);
  });

  it('handles file not found error correctly', async () => {
    const errorMessage = 'File Not Found.';
    requestPostStub.restore();
    sinon.stub(request, 'post').callsFake(async () => {
      throw {
        error: {
          'odata.error': {
            message: {
              lang: 'en-US',
              value: errorMessage
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl
      }
    }), new CommandError(errorMessage));
  });
});
