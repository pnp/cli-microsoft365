import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./sp-get');

describe(commands.SP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const spAppInfo = {
    "value": [
      {
        "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
        "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
        "displayName": "foo",
        "createdDateTime": "2021-03-07T15:04:11Z",
        "description": null,
        "homepage": null,
        "loginUrl": null,
        "logoutUrl": null,
        "notes": null
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.SP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified service principal using its display name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=displayName eq `) > -1) {
        return Promise.resolve(spAppInfo);
      }

      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve(spAppInfo);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, displayName: 'foo' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(spAppInfo));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified service principal using its appId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=appId eq `) > -1) {
        return Promise.resolve(spAppInfo);
      }

      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve(spAppInfo);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, appId: '65415bb1-9267-4313-bbf5-ae259732ee12' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(spAppInfo));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified service principal using its objectId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=objectId eq `) > -1) {
        return Promise.resolve(spAppInfo);
      }

      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve(spAppInfo);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, objectId: '59e617e5-e447-4adc-8b88-00af644d7c92' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(spAppInfo));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no service principal found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/servicePrincipals?api-version=1.6&$filter=displayName eq 'Foo'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ value: [] });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, displayName: 'Foo' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  
  it('fails when Azure AD app with same name exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=displayName eq `) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals",
          "value": [
            {
              "id": "be559819-b036-470f-858b-281c4e808403",
              "appId": "ee091f63-9e48-4697-8462-7cfbf7410b8e",
              "displayName": "foo",
              "createdDateTime": "2021-03-07T15:04:11Z",
              "description": null,
              "homepage": null,
              "loginUrl": null,
              "logoutUrl": null,
              "notes": null
            },
            {
              "id": "93d75ef9-ba9b-4361-9a47-1f6f7478f05f",
              "appId": "e9fd0957-049f-40d0-8d1d-112320fb1cbd",
              "displayName": "foo",
              "createdDateTime": "2021-03-07T15:04:11Z",
              "description": null,
              "homepage": null,
              "loginUrl": null,
              "logoutUrl": null,
              "notes": null
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        displayName: 'foo'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple Azure AD apps with name foo found: be559819-b036-470f-858b-281c4e808403,93d75ef9-ba9b-4361-9a47-1f6f7478f05f`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when the specified Azure AD app does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=displayName eq `) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals",
          "value": []
        });
      }

      return Promise.reject(`Invalid request`);
    });

    command.action(logger, {
      options: {
        debug: true,
        displayName: 'Test App'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified Azure AD app does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither the appId nor the displayName option specified', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = command.validate({ options: { appId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appId option specified', () => {
    const actual = command.validate({ options: { appId: '6a7b1395-d313-4682-8ed4-65a6265a6320' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the displayName option specified', () => {
    const actual = command.validate({ options: { displayName: 'Microsoft Graph' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when both the appId and displayName are specified', () => {
    const actual = command.validate({ options: { appId: '6a7b1395-d313-4682-8ed4-65a6265a6320', displayName: 'Microsoft Graph' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', () => {
    const actual = command.validate({ options: { objectId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and displayName are specified', () => {
    const actual = command.validate({ options: { appId: '123', displayName: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and displayName are specified', () => {
    const actual = command.validate({ options: { displayName: 'abc', objectId: '123' } });
    assert.notStrictEqual(actual, true);
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

  it('supports specifying appId', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying displayName', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});