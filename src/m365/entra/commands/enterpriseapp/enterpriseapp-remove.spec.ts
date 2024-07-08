import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './enterpriseapp-remove.js';
import { settingsNames } from '../../../../settingsNames.js';
import aadCommands from '../../aadCommands.js';

describe(commands.ENTERPRISEAPP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

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

  const deleteRequestStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92') {
        return;
      }

      return new Error('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENTERPRISEAPP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.SP_REMOVE, commands.SP_REMOVE]);
  });

  it('fails when the specified Entra app does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'Invalid'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#servicePrincipals",
          "value": []
        };
      }

      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        appDisplayName: 'Invalid',
        force: true
      }
    }), new CommandError(`The specified Entra app does not exist`));
  });

  it('fails validation if neither the appId nor the appDisplayName option specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    const actual = await command.validate({ options: { appDisplayName: 'Contoso app' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when both the appId and appDisplayName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '6a7b1395-d313-4682-8ed4-65a6265a6320', appDisplayName: 'Contoso app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appObjectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { appObjectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and appDisplayName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '123', appDisplayName: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appDisplayName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appDisplayName: 'abc', appObjectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
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

  it('prompts before removing the enterprise application when force option not passed', async () => {
    await command.action(logger, { options: { appId: '65415bb1-9267-4313-bbf5-ae259732ee12' } });

    assert(promptIssued);
  });

  it('aborts removing the enterprise application when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { appId: '65415bb1-9267-4313-bbf5-ae259732ee12' } });
    assert(deleteCallsSpy.notCalled);
  });

  it('deletes the specified enterprise application using its display name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'foo'`) {
        return spAppInfo;
      }

      throw 'Invalid request';
    });

    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { debug: true, appDisplayName: 'foo', force: true } });
    assert(deleteCallsSpy.calledOnce);
  });

  it('deletes the specified enterprise application using its appId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '65415bb1-9267-4313-bbf5-ae259732ee12'`) {
        return spAppInfo;
      }

      throw 'Invalid request';
    });

    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { appId: '65415bb1-9267-4313-bbf5-ae259732ee12', force: true } });
    assert(deleteCallsSpy.calledOnce);
  });

  it('deletes the specified enterprise application using its appObjectId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=objectId eq '59e617e5-e447-4adc-8b88-00af644d7c92'`) {
        return spAppInfo;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { appObjectId: '59e617e5-e447-4adc-8b88-00af644d7c92', verbose: true } });
    assert(deleteCallsSpy.calledOnce);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails when Entra app with same name exists', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'foo'`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        appDisplayName: 'foo',
        force: true
      }
    }), new CommandError("Multiple Entra apps with name 'foo' found. Found: be559819-b036-470f-858b-281c4e808403, 93d75ef9-ba9b-4361-9a47-1f6f7478f05f."));
  });

  it('handles selecting single result when multiple Entra apps with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'foo'`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(spAppInfo.value[0]);
    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { debug: true, appDisplayName: 'foo', force: true } });
    assert(deleteCallsSpy.calledOnce);
  });
});
