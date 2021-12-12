import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import { copyMock } from './page-copy.mock';
const command: Command = require('./page-copy');

describe(commands.PAGE_COPY, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    Utils.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
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

  it('create a page copy', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(copyMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy (DEBUG)', (done) => {
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

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy and automatically append the aspx extension', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(copyMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy and check if the webUrl is automatically added', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(copyMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy with leading slash in the targetUrl', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/home-copy" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(copyMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy and check if correct URL is used when sitepages is already added', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "sitepages/home-copy" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(copyMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy and check if correct URL is used when sitepages (with leading slash) is already added', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/sitepages/home-copy" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(copyMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy to another site', (done) => {
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

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "https://contoso.sharepoint.com/sites/team-b/sitepages/home-copy" } }, () => {
      try {
        assert(loggerLogSpy.calledWith(copyMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('create a page copy and overwrite the file', (done) => {
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

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('catch any other error in the copy command', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});