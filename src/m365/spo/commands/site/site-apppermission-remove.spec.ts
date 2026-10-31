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
import command, { options } from './site-apppermission-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.SITE_APPPERMISSION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_APPPERMISSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, "'https;//contoso,sharepoint:com/sites/sitecollection-name' is not a valid SharePoint Online site URL.");
  });

  it('passes validation with a correct URL and a filter value', () => {
    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
      appId: '00000000-0000-0000-0000-000000000000'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the appId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
      appId: '123'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, "'123' is not a valid GUID.");
  });

  it('fails validation if appId or appDisplayName or id options are not passed', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if appId, appDisplayName and id options are passed (multiple options)', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
      appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
      appDisplayName: 'Foo',
      id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if appId and appDisplayName both are passed (multiple options)', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
      appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
      appDisplayName: 'Foo'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if appId and id options are passed (multiple options)', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
      appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
      id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if appDisplayName and id options are passed (multiple options)', () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
      appDisplayName: 'Foo',
      id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
    });
    assert.strictEqual(actual.success, false);
  });

  it('prompts before removing the site apppermission when force option not passed', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo'
      })
    });
    assert(promptIssued);
  });

  it('aborts removing the site apppermission when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);

    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo'
      })
    });
    assert(deleteRequestStub.notCalled);
  });

  it('removes site apppermission when prompt confirmed (debug)', async () => {
    sinonUtil.restore(cli.promptForConfirmation);

    sinon.stub(cli, 'promptForConfirmation').resolves(true);

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
      options: commandOptionsSchema.parse({
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        id: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      })
    });
    assert(deleteRequestStub.called);
  });

  it('removes site apppermission with specified appId', async () => {
    sinonUtil.restore(cli.promptForConfirmation);

    sinon.stub(cli, 'promptForConfirmation').resolves(true);

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
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appId: '89ea5c94-7736-4e25-95ad-3fa95f62b66e',
        force: true
      })
    });
    assert(deleteRequestStub.called);
  });

  it('removes site apppermission with specified appDisplayName', async () => {
    sinonUtil.restore(cli.promptForConfirmation);

    sinon.stub(cli, 'promptForConfirmation').resolves(true);

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
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo',
        force: true
      })
    });
    assert(deleteRequestStub.called);
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        appDisplayName: 'Foo',
        force: true
      })
    }), new CommandError('An error has occurred'));
  });
});
