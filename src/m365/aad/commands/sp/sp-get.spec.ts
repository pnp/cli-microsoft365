import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./sp-get');

describe(commands.SP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified service principal using its display name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=displayName eq `) > -1) {
        return Promise.resolve(spAppInfo);
      }

      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve(spAppInfo);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, appDisplayName: 'foo' } });
    assert(loggerLogSpy.calledWith(spAppInfo));
  });

  it('retrieves information about the specified service principal using its appId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=appId eq `) > -1) {
        return Promise.resolve(spAppInfo);
      }

      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve(spAppInfo);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, appId: '65415bb1-9267-4313-bbf5-ae259732ee12' } });
    assert(loggerLogSpy.calledWith(spAppInfo));
  });

  it('retrieves information about the specified service principal using its appObjectId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=objectId eq `) > -1) {
        return Promise.resolve(spAppInfo);
      }

      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals`) > -1) {
        return Promise.resolve(spAppInfo);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, appObjectId: '59e617e5-e447-4adc-8b88-00af644d7c92' } });
    assert(loggerLogSpy.calledWith(spAppInfo));
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

  
  it('fails when Azure AD app with same name exists', async () => {
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

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        appDisplayName: 'foo'
      }
    }), new CommandError(`Multiple Azure AD apps with name foo found: be559819-b036-470f-858b-281c4e808403,93d75ef9-ba9b-4361-9a47-1f6f7478f05f`));
  });

  it('fails when the specified Azure AD app does not exist', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/servicePrincipals?$filter=displayName eq `) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals",
          "value": []
        });
      }

      return Promise.reject(`Invalid request`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        appDisplayName: 'Test App'
      }
    }), new CommandError(`The specified Azure AD app does not exist`));
  });

  it('fails validation if neither the appId nor the appDisplayName option specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appId option specified', async () => {
    const actual = await command.validate({ options: { appId: '6a7b1395-d313-4682-8ed4-65a6265a6320' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the appDisplayName option specified', async () => {
    const actual = await command.validate({ options: { appDisplayName: 'Microsoft Graph' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when both the appId and appDisplayName are specified', async () => {
    const actual = await command.validate({ options: { appId: '6a7b1395-d313-4682-8ed4-65a6265a6320', appDisplayName: 'Microsoft Graph' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appObjectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and appDisplayName are specified', async () => {
    const actual = await command.validate({ options: { appId: '123', appDisplayName: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appDisplayName are specified', async () => {
    const actual = await command.validate({ options: { appDisplayName: 'abc', appObjectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
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

  it('supports specifying appDisplayName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--appDisplayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});