import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./file-copy');

describe(commands.FILE_COPY, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project';
  const documentName = 'Document.pdf';
  const relSourceUrl = '/sites/project/Documents/' + documentName;
  const absoluteSourceUrl = 'https://contoso.sharepoint.com/sites/project/Documents/' + documentName;
  const relTargetUrl = '/sites/IT/Documents';
  const absoluteTargetUrl = 'https://contoso.sharepoint.com/sites/IT/Documents';

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

    requestPostStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/SP.MoveCopyUtil.CopyFileByPath`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
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

  it('fails validation if nameConflictBehavior has an invalid value', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceUrl: absoluteSourceUrl, targetUrl: absoluteTargetUrl, nameConflictBehavior: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required properties are provided', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceUrl: absoluteSourceUrl, targetUrl: absoluteTargetUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('copies file successfully when absolute URLs are provided', async () => {
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
        ShouldBypassSharedLocks: false
      }
    };

    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data, response);
  });

  it('copies file successfully when server relative URLs are provided', async () => {
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
        ShouldBypassSharedLocks: false
      }
    };

    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data, response);
  });

  it('copies file successfully with a new name', async () => {
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
        ShouldBypassSharedLocks: false
      }
    };

    assert.deepStrictEqual(requestPostStub.lastCall.args[0].data, response);
  });

  it('copies file successfully with nameConflictBehavior fail', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        nameConflictBehavior: 'fail'
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.overwrite, false, 'Overwite option is not false');
    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.KeepBoth, false, 'KeepBoth option is not false');
  });

  it('copies file successfully with nameConflictBehavior replace', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        nameConflictBehavior: 'replace'
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.overwrite, true, 'Overwite option is not true');
    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.KeepBoth, false, 'KeepBoth option is not false');
  });

  it('copies file successfully with nameConflictBehavior rename', async () => {
    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        sourceUrl: relSourceUrl,
        targetUrl: relTargetUrl,
        nameConflictBehavior: 'rename'
      }
    });

    assert.strictEqual(requestPostStub.lastCall.args[0].data.overwrite, false, 'Overwite option is not false');
    assert.strictEqual(requestPostStub.lastCall.args[0].data.options.KeepBoth, true, 'KeepBoth option is not true');
  });

  it('copies file successfully with with bypassSharedLock option', async () => {
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
