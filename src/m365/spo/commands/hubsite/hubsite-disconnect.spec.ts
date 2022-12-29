import * as assert from 'assert';
import { AxiosRequestConfig } from 'axios';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
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
const command: Command = require('./hubsite-disconnect');

describe(commands.HUBSITE_DISCONNECT, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const id = '55b979e7-36b6-4968-b3af-6ae221a3483f';
  const title = 'Hub Site';
  const url = 'https://contoso.sharepoint.com/sites/HubSite';

  const etagValue = '1';
  const singleHubSiteResponse = {
    'odata.etag': etagValue,
    ID: id,
    Title: title,
    SiteUrl: url
  };

  const hubSitesResponse = {
    value: [
      singleHubSiteResponse,
      {
        'odata.etag': etagValue,
        ID: 'a9d15b9d-152c-4fa2-be3a-3fbf086f3d49',
        Title: 'Random site title',
        SiteUrl: 'https://contoso.sharepoint.com/sites/RandomSite'
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  let patchStub: sinon.SinonStub<[options: AxiosRequestConfig<any>]>;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      spo.getSpoAdminUrl,
      request.patch
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_DISCONNECT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid id is specified', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if valid title is specified', async () => {
    const actual = await command.validate({ options: { title: title } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if valid url is specified', async () => {
    const actual = await command.validate({ options: { url: url } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before disconnecting the hub site when confirmation argument not passed', async () => {
    await command.action(logger, { options: { id: id } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts disconnecting hub site when prompt not confirmed', async () => {
    await command.action(logger, { options: { url: url } });
    assert(patchStub.notCalled);
  });

  it('disconnects hub site when when id option is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites/GetById('${id}')`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return singleHubSiteResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        id: id,
        verbose: true,
        confirm: true
      }
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: '00000000-0000-0000-0000-000000000000' }, 'Request body does not match');
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue, 'if-match request header doesn\'t match');
  });

  it('disconnects hub site when when title option is specified', async () => {
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
        title: title,
        verbose: true,
        confirm: true
      }
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: '00000000-0000-0000-0000-000000000000' }, 'Request body does not match');
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue, 'if-match request header doesn\'t match');
  });

  it('disconnects hub site when when url option is specified', async () => {
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
        url: url,
        verbose: true,
        confirm: true
      }
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { ParentHubSiteId: '00000000-0000-0000-0000-000000000000' }, 'Request body does not match');
    assert.deepStrictEqual(patchStub.lastCall.args[0].headers!['if-match'], etagValue, 'if-match request header doesn\'t match');
  });

  it('disconnects hub site when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return hubSitesResponse;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        title: title
      }
    });
  });

  it('throws an error when multiple hub sites with the same title were retrieved', async () => {
    const response = {
      value: [
        singleHubSiteResponse,
        {
          'odata.etag': etagValue,
          ID: 'a9d15b9d-152c-4fa2-be3a-3fbf086f3d49',
          Title: title,
          SiteUrl: 'https://contoso.sharepoint.com/sites/RandomSite'
        }
      ]
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return response;
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        confirm: true
      }
    }), new CommandError(`Multiple hub sites with name '${title}' found: ${response.value.map(s => s.ID).join(',')}.`));
  });

  it('throws an error when no hub sites with the same title were found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return {
            value: []
          };
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        confirm: true
      }
    }), new CommandError(`The specified hub site '${title}' does not exist.`));
  });

  it('throws an error when no hub sites with the same url were found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const baseUrl = opts.url?.split('?')[0];
      if (baseUrl === `${spoAdminUrl}/_api/HubSites`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json;odata=minimalmetadata') !== -1) {
          return {
            value: []
          };
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: {
        url: url,
        confirm: true
      }
    }), new CommandError(`The specified hub site '${url}' does not exist.`));
  });

  it('handles random API error correctly', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_api/HubSites/GetById('${id}')?$select=ID`) {
        return singleHubSiteResponse;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    const errorMessage = 'Something went wrong';
    patchStub.restore();
    sinon.stub(request, 'patch').callsFake(async () => { throw { error: { 'odata.error': { message: { value: errorMessage } } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        confirm: true
      }
    }), new CommandError(errorMessage));
  });
});
