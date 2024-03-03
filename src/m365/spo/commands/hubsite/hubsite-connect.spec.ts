import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './hubsite-connect.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.HUBSITE_CONNECT, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const id = '55b979e7-36b6-4968-b3af-6ae221a3483f';
  const parentId = 'f7510a39-8423-43fd-aed8-e3b11d043e0f';
  const title = 'Hub Site';
  const parentTitle = 'Parent Hub Site';
  const url = 'https://contoso.sharepoint.com/sites/HubSite';
  const parentUrl = 'https://contoso.sharepoint.com/sites/ParentHubSite';

  const etagValue = '1';
  const hubSitesResponse = {
    value: [
      {
        'odata.etag': etagValue,
        ID: id,
        Title: title,
        SiteUrl: url
      },
      {
        'odata.etag': '3',
        ID: parentId,
        Title: parentTitle,
        SiteUrl: parentUrl
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let patchStub: sinon.SinonStub<[options: CliRequestOptions]>;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);

    sinon.stub(spo, 'getSpoAdminUrl').resolves(spoAdminUrl);
    patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/HubSites/GetById('${id}')`) {
        return;
      }

      throw 'Invalid requet URL: ' + opts.url;
    });
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
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HUBSITE_CONNECT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'Invalid', parentTitle: parentTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if parentId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { title: title, parentId: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'Invalid', parentTitle: parentTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if parentUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { title: title, parentUrl: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid id and parentId are specified', async () => {
    const actual = await command.validate({ options: { id: id, parentId: parentId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if title and parentTitle are specified', async () => {
    const actual = await command.validate({ options: { title: title, parentTitle: parentTitle } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if url and parentUrl are specified', async () => {
    const actual = await command.validate({ options: { url: url, parentUrl: parentUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly connects hub site to parent hub site by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return hubSitesResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: id,
        parentId: parentId
      }
    });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: parentId });
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue);
  });

  it('correctly connects hub site to parent hub site by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return hubSitesResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        title: title,
        parentTitle: parentTitle
      }
    });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: parentId });
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue);
  });

  it('correctly connects hub site to parent hub site by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return hubSitesResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        url: url,
        parentUrl: parentUrl
      }
    });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: parentId });
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue);
  });

  it('throws error when hub site with ID was not found', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        parentId: parentId
      }
    }), new CommandError(`The specified hub site '${id}' does not exist.`));
  });

  it('throws error when hub site with title was not found', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        parentTitle: parentTitle
      }
    }), new CommandError(`The specified hub site '${title}' does not exist.`));
  });

  it('throws error when hub site with url was not found', async () => {
    sinon.stub(request, 'get').resolves({ value: [] });

    await assert.rejects(command.action(logger, {
      options: {
        url: url,
        parentUrl: parentUrl
      }
    }), new CommandError(`The specified hub site '${url}' does not exist.`));
  });

  it('throws error when multiple hub sites with the same name were found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').resolves({
      value: [
        {
          Title: title,
          ID: id
        },
        {
          Title: title,
          ID: parentId
        }
      ]
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        parentUrl: parentUrl
      }
    }), new CommandError("Multiple hub sites with name 'Hub Site' found. Found: 55b979e7-36b6-4968-b3af-6ae221a3483f, f7510a39-8423-43fd-aed8-e3b11d043e0f."));
  });

  it('handles selecting single result when multiple hubsites with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return {
            value: [
              {
                Title: title,
                ID: id
              },
              {
                Title: title,
                ID: parentId
              },
              {
                Title: parentTitle,
                ID: id
              },
              {
                Title: parentTitle,
                ID: parentId
              }
            ]
          };
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({
      Title: title,
      ID: id,
      'odata.etag': etagValue
    });

    await command.action(logger, {
      options: {
        verbose: true,
        title: title,
        parentTitle: parentTitle
      }
    });
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue);
  });

  it('correctly handles random error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          message: {
            value: 'Something went wrong'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        parentUrl: parentUrl
      }
    }), new CommandError('Something went wrong'));
  });
});
