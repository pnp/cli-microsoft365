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

  it('gets information about the specified hub site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites/getbyid('ee8b42c3-3e6f-4822-87c1-c21ad666046b')`) > -1) {
        return Promise.resolve({
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: 'ee8b42c3-3e6f-4822-87c1-c21ad666046b' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites/getbyid('ee8b42c3-3e6f-4822-87c1-c21ad666046b')`) > -1) {
        return Promise.resolve({
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: 'ee8b42c3-3e6f-4822-87c1-c21ad666046b' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        }));
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
              "value": "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, id: 'ee8b42c3-3e6f-4822-87c1-c21ad666046b' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
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
});