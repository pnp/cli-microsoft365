import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./hubsite-get');

describe(commands.HUBSITE_GET, () => {
  const validId = '9ff01368-1183-4cbb-82f2-92e7e9a3f4ce';
  const validTitle = 'Hub Site';
  const validUrl = 'https://contoso.sharepoint.com';

  const hubsiteResponse = {
    "ID": validId,
    "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
    "SiteUrl": validUrl,
    "Title": validTitle
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets();
    assert.deepStrictEqual(optionSets, [['id', 'title', 'url']]);
  });

  it('gets information about the specified hub site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites/getbyid('${validId}')`) > -1) {
        return Promise.resolve(hubsiteResponse);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: validId } }, () => {
      try {
        assert(loggerLogSpy.calledWith(hubsiteResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites/getbyid('${validId}')`) > -1) {
        return Promise.resolve(hubsiteResponse);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: validId } }, () => {
      try {
        assert(loggerLogSpy.calledWith(hubsiteResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site by title', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({ value: [hubsiteResponse] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: validTitle } }, () => {
      try {
        assert(loggerLogSpy.calledWith(hubsiteResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site by url', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({ value: [hubsiteResponse] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: validUrl } }, () => {
      try {
        assert(loggerLogSpy.calledWith(hubsiteResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple hubsites found with same title', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({ value: [hubsiteResponse, hubsiteResponse] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: validTitle } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple hub sites with ${validTitle} found. Please disambiguate: ${validUrl}, ${validUrl}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when no hubsites found with title', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: validTitle } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified hub site ${validTitle} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple hubsites found with same url', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({ value: [hubsiteResponse, hubsiteResponse] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: validUrl } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple hub sites with ${validUrl} found. Please disambiguate: ${validUrl}, ${validUrl}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when no hubsites found with url', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: validUrl } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified hub site ${validUrl} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when hub site not found', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
            "message": {
              "lang": "en-US",
              "value": "The specified hub site with id ee8b42c3-3e6f-4822-87c1-c21ad666046b does not exist"
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, id: 'ee8b42c3-3e6f-4822-87c1-c21ad666046b' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified hub site with id ee8b42c3-3e6f-4822-87c1-c21ad666046b does not exist`)));
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

  it('supports specifying id', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = command.validate({ options: { id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', () => {
    const actual = command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } });
    assert.strictEqual(actual, true);
  });

  it(`fails validation if the specified url is invalid`, () => {
    const actual = command.validate({ options: {
      url: '/'
    } });
    assert.notStrictEqual(actual, true);
  });
});