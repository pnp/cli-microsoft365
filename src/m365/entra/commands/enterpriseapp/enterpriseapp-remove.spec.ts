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
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/59e617e5-e447-4adc-8b88-00af644d7c92` ||
        opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals(appId='65415bb1-9267-4313-bbf5-ae259732ee12')`
      ) {
        return;
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);
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

    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation,
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

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [commands.SP_REMOVE]);
  });

  it('fails when the specified enterprise applications does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'Invalid'&$select=id`) {
        return {
          value: []
        };
      }

      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        displayName: 'Invalid',
        force: true
      }
    }), new CommandError(`The specified enterprise application does not exist.`));
  });

  it('fails validation if neither the id nor the displayName option is specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id option is specified', async () => {
    const actual = await command.validate({ options: { id: '6a7b1395-d313-4682-8ed4-65a6265a6320' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the displayName option is specified', async () => {
    const actual = await command.validate({ options: { displayName: 'Contoso app' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when both the id and displayName are specified', async () => {
    const actual = await command.validate({ options: { id: '6a7b1395-d313-4682-8ed4-65a6265a6320', displayName: 'Contoso app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the objectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { objectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and displayName are specified', async () => {
    const actual = await command.validate({ options: { id: '123', displayName: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and displayName are specified', async () => {
    const actual = await command.validate({ options: { displayName: 'abc', objectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('prompts before removing the enterprise application when force option not passed', async () => {
    await command.action(logger, { options: { id: '65415bb1-9267-4313-bbf5-ae259732ee12' } });

    assert(promptIssued);
  });

  it('aborts removing the enterprise application when prompt not confirmed', async () => {
    const deleteCallsSpy = sinon.stub(request, 'delete').resolves();
    await command.action(logger, { options: { id: '65415bb1-9267-4313-bbf5-ae259732ee12' } });
    assert(deleteCallsSpy.notCalled);
  });

  it('deletes the specified enterprise application using its display name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'foo'&$select=id`) {
        return spAppInfo;
      }

      throw 'Invalid request';
    });

    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { verbose: true, displayName: 'foo', force: true } });
    assert(deleteCallsSpy.calledOnce);
  });

  it('deletes the specified enterprise application using its id', async () => {
    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { id: '65415bb1-9267-4313-bbf5-ae259732ee12', force: true } });
    assert(deleteCallsSpy.calledOnce);
  });

  it('deletes the specified enterprise application using its objectId', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const deleteCallsSpy: sinon.SinonStub = deleteRequestStub();
    await command.action(logger, { options: { objectId: '59e617e5-e447-4adc-8b88-00af644d7c92', verbose: true } });
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

    await assert.rejects(command.action(logger, { options: { displayName: 'foo', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails when enterprise applications with same name exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'foo'&$select=id`) {
        return {
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
        verbose: true,
        displayName: 'foo',
        force: true
      }
    }), new CommandError("Multiple enterprise applications with name 'foo' found. Found: be559819-b036-470f-858b-281c4e808403, 93d75ef9-ba9b-4361-9a47-1f6f7478f05f."));
  });

  it('handles selecting single result when multiple enterprise applications with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq 'foo'&$select=id`) {
        return {
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
    await command.action(logger, { options: { verbose: true, displayName: 'foo', force: true } });
    assert(deleteCallsSpy.calledOnce);
  });
});
