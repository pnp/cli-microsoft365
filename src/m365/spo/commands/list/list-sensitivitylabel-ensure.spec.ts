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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-sensitivitylabel-ensure');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_SENSITIVITYLABEL_ENSURE), true);
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

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: 'invalid', name: name } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, name: name } }, commandInfo);
    assert(actual);
  });

  it('fails validation if listId, listUrl and listTitle options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, name: name } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should apply sensitivity label to document library using title', async () => {
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels?$filter=name eq '${name}'&$select=id`) {
        return graphResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        return Promise.reject(new Error("404 - \"404 FILE NOT FOUND\""));
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        name: name
      }
    } as any), new CommandError('404 - "404 FILE NOT FOUND"'));
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

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('Shared%20Documents')`) {
        return Promise.reject({
          error: {
            'odata.error': {
              code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
              message: {
                value: 'Can not find compliance tag with value: abc. SiteSubscriptionId: ea1787c6-7ce2-4e71-be47-5e0deb30f9e4'
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
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