import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./page-remove');

describe(commands.PAGE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

  const fakeRestCalls: (pageName?: string) => sinon.SinonStub = (pageName: string = 'page.aspx') => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/${pageName}')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon
      .stub(spo, 'getRequestDigest')
      .callsFake(() => Promise.resolve({
        FormDigestValue: 'ABC',
        FormDigestTimeoutSeconds: 1800,
        FormDigestExpiresAt: new Date(),
        WebFullUrl: 'https://contoso.sharepoint.com'
      }));
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
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
      spo.getRequestDigest,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
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
          debug: false,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          confirm: true
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
          confirm: true
        }
      });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes a modern page (debug) without confirm prompt on root of tenant', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sitepages/page.aspx')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com',
          confirm: true
        }
      });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes a modern page with confirm prompt', async () => {
    fakeRestCalls();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: true };
    });
    await command.action(logger,
      {
        options: {
          debug: false,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('removes a modern page (debug) with confirm prompt', async () => {
    fakeRestCalls();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: true };
    });
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
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('should abort page removal when prompt not confirmed', async () => {
    const postCallSpy = fakeRestCalls();
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
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
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger,
      {
        options: {
          debug: false,
          name: 'page',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          confirm: true
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when removing modern page', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await assert.rejects(command.action(logger,
      {
        options: {
          debug: false,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          confirm: true
        }
      }), new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
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
      if (o.option.indexOf('--confirm') > -1) {
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
