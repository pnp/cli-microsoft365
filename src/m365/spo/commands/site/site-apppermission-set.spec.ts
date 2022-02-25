import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./site-apppermission-set');

describe(commands.SITE_APPPERMISSION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_APPPERMISSION_SET), true);
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

  it('fails validation if permissionId, appId, and appDisplayName options are not specified', () => {
    const actual = command.validate({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write"
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
    };
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

  it('fails to get Azure AD app when Azure AD app does not exists', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/permissions') > -1) {
          return Promise.resolve({ value: [] });
        }
        return Promise.reject('The specified app permission does not exist');
      });

    command.action(logger, {
      options: {
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified app permission does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple Azure AD apps with same name exists', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Multiple app permissions with displayName Foo found: 89ea5c94-7736-4e25-95ad-3fa95f62b66e,cca00169-d38b-462f-a3b4-f3566b162f2d7');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/permissions') > -1) {
          return Promise.resolve({
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
                "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                "grantedToIdentities": [
                  {
                    "application": {
                      "displayName": "Foo",
                      "id": "cca00169-d38b-462f-a3b4-f3566b162f2d7"
                    }
                  }
                ]
              }
            ]
          });
        }
        return Promise.reject('Multiple app permissions with displayName Foo found: 89ea5c94-7736-4e25-95ad-3fa95f62b66e,cca00169-d38b-462f-a3b4-f3566b162f2d7');
      });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appDisplayName: "Foo"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple app permissions with displayName Foo found: 89ea5c94-7736-4e25-95ad-3fa95f62b66e,cca00169-d38b-462f-a3b4-f3566b162f2d7`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Updates an application permission to the site by appId', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/permissions') > -1) {
          return Promise.resolve({
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
                "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                "grantedToIdentities": [
                  {
                    "application": {
                      "displayName": "TeamsBotDemo5",
                      "id": "cca00169-d38b-462f-a3b4-f3566b162f2d"
                    }
                  }
                ]
              }
            ]
          });
        }

        return Promise.reject('Invalid request');
      });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return Promise.resolve({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        output: "json"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Updates an application permission to the site by appDisplayName', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/permissions') > -1) {
          return Promise.resolve({
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
                "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                "grantedToIdentities": [
                  {
                    "application": {
                      "displayName": "TeamsBotDemo5",
                      "id": "cca00169-d38b-462f-a3b4-f3566b162f2d"
                    }
                  }
                ]
              }
            ]
          });
        }

        return Promise.reject('Invalid request');
      });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return Promise.resolve({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appDisplayName: "Foo",
        output: "json"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Updates an application permission to the site by permissionId', (done) => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return Promise.resolve({
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          });
        }
        return Promise.reject('Invalid request');
      });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return Promise.resolve({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        permissionId: "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
        output: "json"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
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
