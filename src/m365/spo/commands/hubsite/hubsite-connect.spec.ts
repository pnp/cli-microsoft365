import * as assert from 'assert';
import { AxiosRequestConfig } from 'axios';
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
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./hubsite-connect');

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
  let patchStub: sinon.SinonStub<[options: AxiosRequestConfig<any>]>;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);

    sinon.stub(spo, 'getSpoAdminUrl').callsFake(async () => spoAdminUrl);
    patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/HubSites/GetById('${id}')`) {
        return;
      }

      throw 'Invalid requet URL: ' + opts.url;
    });
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName,
      spo.getSpoAdminUrl,
      request.patch
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_CONNECT), true);
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
    sinon.stub(request, 'get').callsFake(async () => ({ value: [] }));

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        parentId: parentId
      }
    }), new CommandError(`The specified hub site '${id}' does not exist.`));
  });

  it('throws error when hub site with title was not found', async () => {
    sinon.stub(request, 'get').callsFake(async () => ({ value: [] }));

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        parentTitle: parentTitle
      }
    }), new CommandError(`The specified hub site '${title}' does not exist.`));
  });

  it('throws error when hub site with url was not found', async () => {
    sinon.stub(request, 'get').callsFake(async () => ({ value: [] }));

    await assert.rejects(command.action(logger, {
      options: {
        url: url,
        parentUrl: parentUrl
      }
    }), new CommandError(`The specified hub site '${url}' does not exist.`));
  });

  it('throws error when multiple hub sites with the same name were found', async () => {
    sinon.stub(request, 'get').callsFake(async () => ({
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
    }));

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        parentUrl: parentUrl
      }
    }), new CommandError(`Multiple hub sites with name '${title}' found: ${id},${parentId}.`));
  });

  it('correctly handles random error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw {
        error: {
          'odata.error': {
            message: {
              value: 'Something went wrong'
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        parentUrl: parentUrl
      }
    }), new CommandError('Something went wrong'));
  });
});