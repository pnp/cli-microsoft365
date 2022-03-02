import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./sitedesign-rights-list');

describe(commands.SITEDESIGN_RIGHTS_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_RIGHTS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about permissions granted for the specified site design', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "value": [
            {
              "DisplayName": "MOD Administrator",
              "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
              "Rights": "1"
            },
            {
              "DisplayName": "Patti Fernandez",
              "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
              "Rights": "1"
            }
          ]
        }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "DisplayName": "MOD Administrator",
            "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Rights": "View"
          },
          {
            "DisplayName": "Patti Fernandez",
            "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
            "Rights": "View"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about permissions granted for the specified site design (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "value": [
            {
              "DisplayName": "MOD Administrator",
              "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
              "Rights": "1"
            },
            {
              "DisplayName": "Patti Fernandez",
              "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
              "Rights": "1"
            }
          ]
        }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "DisplayName": "MOD Administrator",
            "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Rights": "View"
          },
          {
            "DisplayName": "Patti Fernandez",
            "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
            "Rights": "View"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns original value for unknown permissions', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "value": [
            {
              "DisplayName": "MOD Administrator",
              "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
              "Rights": "1"
            },
            {
              "DisplayName": "Patti Fernandez",
              "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
              "Rights": "2"
            }
          ]
        }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "DisplayName": "MOD Administrator",
            "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Rights": "View"
          },
          {
            "DisplayName": "Patti Fernandez",
            "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
            "Rights": "2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when site script not found', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });
    });

    command.action(logger, { options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
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