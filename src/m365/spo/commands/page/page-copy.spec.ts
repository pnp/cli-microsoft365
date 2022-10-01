import * as assert from 'assert';
import * as chalk from 'chalk';
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
import commands from '../../commands';
import { copyMock } from './page-copy.mock';
const command: Command = require('./page-copy');

describe(commands.PAGE_COPY, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.PAGE_COPY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'PageLayoutType', 'Title', 'Url']);
  });

  it('create a page copy', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy (DEBUG)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } });
    assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
  });

  it('create a page copy and automatically append the aspx extension', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and check if the webUrl is automatically added', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy with leading slash in the targetUrl', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and check if correct URL is used when sitepages is already added', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "sitepages/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and check if correct URL is used when sitepages (with leading slash) is already added', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/sitepages/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy to another site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-b/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "https://contoso.sharepoint.com/sites/team-b/sitepages/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and overwrite the file', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return Promise.resolve(copyMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } });
    assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
  });

  it('catch any other error in the copy command', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } }), new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', sourceName: 'home.aspx', targetUrl: 'home-copy.aspx' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', sourceName: 'home.aspx', targetUrl: 'home-copy.aspx' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});