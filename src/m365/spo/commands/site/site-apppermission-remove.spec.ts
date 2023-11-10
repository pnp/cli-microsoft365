import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-apppermission-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.SITE_APPPERMISSION_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

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
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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

    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;

    deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/permissions/') > -1) {
        return;
      }
      throw '';
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      global.setTimeout,
      Cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_APPPERMISSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with a correct URL and a filter value', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '123'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId or appDisplayName or id options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId, appDisplayName and id options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        appDisplayName: 'Foo',
        id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appDisplayName both are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        appDisplayName: 'Foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and id options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appDisplayName and id options are passed (multiple options)', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo',
        id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('prompts before removing the site apppermission when force option not passed', async () => {
    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo'
      }
    });
    assert(promptIssued);
  });

  it('aborts removing the site apppermission when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);

    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo'
      }
    });
    assert(deleteRequestStub.notCalled);
  });

  it('removes site apppermission when prompt confirmed (debug)', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);

    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf(":/sites/sitecollection-name") > - 1) {
          return site;
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf("contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000/permissions") > - 1) {
          return response;
        }
        throw 'Invalid request';
      });

    await command.action(logger, {
      options: {
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    });
    assert(deleteRequestStub.called);
  });

  it('removes site apppermission with specified appId', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);

    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf(":/sites/sitecollection-name") > - 1) {
          return site;
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf("contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000/permissions") > - 1) {
          return response;
        }
        throw 'Invalid request';
      });

    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        force: true
      }
    });
    assert(deleteRequestStub.called);
  });

  it('removes site apppermission with specified appDisplayName', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);

    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf(":/sites/sitecollection-name") > - 1) {
          return site;
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf("contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000/permissions") > - 1) {
          return response;
        }
        throw 'Invalid request';
      });

    await command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo',
        force: true
      }
    });
    assert(deleteRequestStub.called);
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo',
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
