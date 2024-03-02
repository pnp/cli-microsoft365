import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-sensitivitylabel-ensure.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LIST_SENSITIVITYLABEL_ENSURE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const name = 'Label';
  const listTitle = 'Shared Documents';
  const listId = 'b4cfa0d9-b3d7-49ae-a0f0-f14ffdd005f7';
  const listUrl = '/Shared Documents';
  const sensitivityLabelId = '0d39dc11-75ff-4309-8b32-ff94f0e41607';
  const graphResponse = {
    "value": [
      {
        "id": sensitivityLabelId,
        "name": "Label",
        "description": "",
        "color": "",
        "sensitivity": 7,
        "tooltip": "sensitive information.",
        "isActive": true,
        "isAppliable": true
      }
    ]
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_SENSITIVITYLABEL_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: listId, name: name } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, name: name } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, id: sensitivityLabelId } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: 'invalid', name: name } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, name: name } }, commandInfo);
    assert(actual);
  });

  it('fails validation if id and name options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId, listUrl and listTitle options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, name: name } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should apply sensitivity label by id to document library using title', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, id: sensitivityLabelId, verbose: true } } as any);

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, sensitivityLabelId);
  });

  it('should apply sensitivity label by name to document library using title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${name}'&$select=id`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, name: name, verbose: true } } as any);

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, sensitivityLabelId);
  });

  it('should apply sensitivity label to document library using URL', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${name}'&$select=id`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList('%2FShared%20Documents')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listUrl: listUrl, name: name, verbose: true } } as any);

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, sensitivityLabelId);
  });

  it('should apply sensitivity label to document library using id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${name}'&$select=id`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'b4cfa0d9-b3d7-49ae-a0f0-f14ffdd005f7')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listId: listId, name: name, verbose: true } } as any);

    const lastCall = patchStub.lastCall.args[0];
    assert.strictEqual(lastCall.data.DefaultSensitivityLabelForLibrary, sensitivityLabelId);
  });

  it('should handle error if list does not exist', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: '404 - File not found'
          }
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${name}'&$select=id`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        name: name
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });

  it('should handle error if the specified sensitivity label does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${name}'&$select=id`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        name: name
      }
    } as any), new CommandError('The specified sensitivity label does not exist'));
  });

  it('should handle error when trying to set label', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${name}'&$select=id`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        throw {
          error: {
            'odata.error': {
              code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
              message: {
                value: 'Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        name: name
      }
    } as any), new CommandError("Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"));
  });
});