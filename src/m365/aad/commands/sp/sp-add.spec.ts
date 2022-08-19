import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./sp-add');

describe(commands.SP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SP_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if neither the appId, appName, nor objectId option specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { objectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and appName are specified', async () => {
    const actual = await command.validate({ options: { appId: '00000000-0000-0000-0000-000000000000', appName: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appName and objectId are specified', async () => {
    const actual = await command.validate({ options: { appName: 'abc', objectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and objectId are specified', async () => {
    const actual = await command.validate({ options: { appId: '00000000-0000-0000-0000-000000000000', objectId: '00000000-0000-0000-0000-000000000000' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appId option specified', async () => {
    const actual = await command.validate({ options: { appId: '00000000-0000-0000-0000-000000000000' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the appName option specified', async () => {
    const actual = await command.validate({ options: { appName: 'abc' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the objectId option specified', async () => {
    const actual = await command.validate({ options: { objectId: '00000000-0000-0000-0000-000000000000' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying appId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying appName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying objectId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--objectId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('correctly handles API OData error', async () => {
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

    await assert.rejects(command.action(logger, { options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles no service principal found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({ value: [] });
        }
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, appName: 'Foo' } });
    assert(loggerLogSpy.notCalled);
  });

  it('fails when the specified Azure AD app does not exist', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/applications?$filter=id eq `) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
          "value": [
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        objectId: '59e617e5-e447-4adc-8b88-00af644d7c92'
      }
    }), new CommandError(`The specified Azure AD app doesn't exist`));
  });

  it('fails when Azure AD app with same name exists', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/applications?$filter=displayName eq `) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications",
          "value": [
            {
              "id": "be559819-b036-470f-858b-281c4e808403",
              "appId": "ee091f63-9e48-4697-8462-7cfbf7410b8e",
              "displayName": "Test App"
            },
            {
              "id": "93d75ef9-ba9b-4361-9a47-1f6f7478f05f",
              "appId": "e9fd0957-049f-40d0-8d1d-112320fb1cbd",
              "displayName": "Test App"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        appName: 'Test App'
      }
    }), new CommandError(`Multiple Azure AD apps with name Test App found: ee091f63-9e48-4697-8462-7cfbf7410b8e,e9fd0957-049f-40d0-8d1d-112320fb1cbd`));
  });

  it('adds a service principal to a registered Azure AD app by appId', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals`) {
        return Promise.resolve({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        });
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        appId: '65415bb1-9267-4313-bbf5-ae259732ee12'
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
      "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
      "displayName": "foo"
    }));
  });

  it('adds a service principal to a registered Azure AD app by appName', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/applications?$filter=displayName eq `) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
              "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
              "displayName": "foo"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        appName: 'foo'
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
      "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
      "displayName": "foo"
    }));
  });

  it('adds a service principal to a registered Azure AD app by objectId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/applications?$filter=id eq `) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
              "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
              "displayName": "foo"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve({
          "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
          "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
          "displayName": "foo"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        objectId: '59e617e5-e447-4adc-8b88-00af644d7c92'
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "59e617e5-e447-4adc-8b88-00af644d7c92",
      "appId": "65415bb1-9267-4313-bbf5-ae259732ee12",
      "displayName": "foo"
    }));
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
});
