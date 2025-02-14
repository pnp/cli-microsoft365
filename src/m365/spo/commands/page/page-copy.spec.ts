import assert from 'assert';
import chalk from 'chalk';
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
import command from './page-copy.js';
import { copyMock } from './page-copy.mock.js';

describe(commands.PAGE_COPY, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_COPY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'PageLayoutType', 'Title', 'Url']);
  });

  it('create a page copy', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy (DEBUG)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } });
    assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
  });

  it('create a page copy and automatically append the aspx extension', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and check if the webUrl is automatically added', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy with leading slash in the targetUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and check if correct URL is used when sitepages is already added', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "sitepages/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and check if correct URL is used when sitepages (with leading slash) is already added', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/sitepages/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy to another site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-b/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "https://contoso.sharepoint.com/sites/team-b/sitepages/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and overwrite the file', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } });
    assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
  });

  it('catch any other error in the copy command', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } }), new CommandError('An error has occurred'));
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
