import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./site-apppermission-remove');

describe(commands.SITE_APPPERMISSION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  let deleteRequestStub: sinon.SinonStub;

  const site = {
    "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
    "displayName": "OneDrive Team Site",
    "name": "1drvteam",
    "createdDateTime": "2017-05-09T20:56:00Z",
    "lastModifiedDateTime": "2017-05-09T20:56:01Z",
    "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
  };

  const response = {
    "value": [
      {
        "id": "aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
        "grantedToIdentities": [
          {
            "application": {
              "displayName": "Foo",
              "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
            }
          }
        ]
      },
      {
        "id": "aTowaS50fG1zLnNwLmV4dHxkMDVhMmRkYi0xZjMzLTRkZTMtOTMzNS0zYmZiZTUwNDExYzVAZWUyYjdjMGMtZDI1My00YjI3LTk0NmItMDYzZGM4OWNlOGMy",
        "grantedToIdentities": [
          {
            "application": {
              "displayName": "SPRestSample",
              "id": "d05a2ddb-1f33-4de3-9335-3bfbe50411c5"
            }
          }
        ]
      }
    ]
  };

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

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });

    promptOptions = undefined;

    deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions/') > -1) {
        return Promise.resolve();
      }
      return Promise.reject();
    });
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.delete,
      global.setTimeout,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.SITE_APPPERMISSION_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', (done) => {
    const actual = command.validate({
      options: {
        siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation with a correct URL and a filter value', (done) => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '00000000-0000-0000-0000-000000000000'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '123'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId or appDisplayName or permissionId options are not passed', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId, appDisplayName and permissionId options are passed (multiple options)', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        appDisplayName: 'Foo',
        permissionId: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appDisplayName both are passed (multiple options)', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        appDisplayName: 'Foo'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and permissionId options are passed (multiple options)', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        permissionId: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appDisplayName and permissionId options are passed (multiple options)', () => {
    const actual = command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo',
        permissionId: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('prompts before removing the site apppermission when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo'
      }
    }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the site apppermission when prompt not confirmed', (done) => {
    Utils.restore(Cli.prompt);

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo'
      }
    }, () => {
      try {
        assert(deleteRequestStub.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site apppermission when prompt confirmed (debug)', (done) => {
    Utils.restore(Cli.prompt);

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf(":/sites/sitecollection-name") > - 1) {
          return Promise.resolve(site);
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf("contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000/permissions") > - 1) {
          return Promise.resolve(response);
        }
        return Promise.reject('Invalid request');
      });

    command.action(logger, {
      options: {
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permissionId: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site apppermission with specified appId', (done) => {
    Utils.restore(Cli.prompt);

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf(":/sites/sitecollection-name") > - 1) {
          return Promise.resolve(site);
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf("contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000/permissions") > - 1) {
          return Promise.resolve(response);
        }
        return Promise.reject('Invalid request');
      });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        confirm: true
      }
    }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes site apppermission with specified appDisplayName', (done) => {
    Utils.restore(Cli.prompt);

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf(":/sites/sitecollection-name") > - 1) {
          return Promise.resolve(site);
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf("contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000/permissions") > - 1) {
          return Promise.resolve(response);
        }
        return Promise.reject('Invalid request');
      });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo',
        confirm: true
      }
    }, () => {
      try {
        assert(deleteRequestStub.called);
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
