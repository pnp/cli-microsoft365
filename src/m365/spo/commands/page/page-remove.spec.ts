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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './page-remove.js';

describe(commands.PAGE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptIssued: boolean = false;

  const fakeRestCalls: (pageName?: string) => sinon.SinonStub = (pageName: string = 'page.aspx') => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/${pageName}')`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon
      .stub(spo, 'getRequestDigest').resolves({
        FormDigestValue: 'ABC',
        FormDigestTimeoutSeconds: 1800,
        FormDigestExpiresAt: new Date(),
        WebFullUrl: 'https://contoso.sharepoint.com'
      });
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
    loggerLogSpy = sinon.spy(logger, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.PAGE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a modern page without confirm prompt', async () => {
    fakeRestCalls();
    await command.action(logger,
      {
        options: {
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('removes a modern page (debug) without confirm prompt', async () => {
    fakeRestCalls();
    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes a modern page (debug) without confirm prompt on root of tenant', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sitepages/page.aspx')`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com',
          force: true
        }
      });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes a modern page with confirm prompt', async () => {
    fakeRestCalls();
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger,
      {
        options: {
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('removes a modern page (debug) with confirm prompt', async () => {
    fakeRestCalls();
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });
    assert(loggerLogToStderrSpy.called);
  });

  it('should prompt before removing page when confirmation argument not passed', async () => {
    fakeRestCalls();
    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });

    assert(promptIssued);
  });

  it('should abort page removal when prompt not confirmed', async () => {
    const postCallSpy = fakeRestCalls();
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });
    assert(postCallSpy.notCalled === true);
  });

  it('automatically appends the .aspx extension', async () => {
    fakeRestCalls();
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger,
      {
        options: {
          name: 'page',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when removing modern page', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await assert.rejects(command.action(logger,
      {
        options: {
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      }), new CommandError('An error has occurred'));
  });

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirm', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--force') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({
      options: { name: 'page.aspx', webUrl: 'http://foo' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({
      options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', async () => {
    const actual = await command.validate({
      options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
