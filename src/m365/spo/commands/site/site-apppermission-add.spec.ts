import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./site-apppermission-add');

describe(commands.SITE_APPPERMISSION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.SITE_APPPERMISSION_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', (done) => {
    const actual = command.validate({
      options: {
        siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = command.validate({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "123",
        appDisplayName: "Foo App"
      }
    });

    assert.notStrictEqual(actual, true);
  });

  it('passes validation with a correct URL', (done) => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails validation if invalid value specified for permission', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "Invalid",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails when passing a site that does not exist', (done) => {
    const siteError = {
      "error": {
        "code": "itemNotFound",
        "message": "Requested site could not be found",
        "innerError": {
          "date": "2021-03-03T08:58:02",
          "request-id": "4e054f93-0eba-4743-be47-ce36b5f91120",
          "client-request-id": "dbd35b28-0ec3-6496-1279-0e1da3d028fe"
        }
      }
    }
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('non-existing') === -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject(siteError);
    });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name-non-existing',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Requested site could not be found")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Adds an application permission to the site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return Promise.resolve({
          "roles": ["write"],
          "grantedToIdentities": [{
            "application": { "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e", "displayName": "Foo App" }
          }]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "roles": ["write"],
          "grantedToIdentities": [{
            "application": { "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e", "displayName": "Foo App" }
          }]
        }));
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
});